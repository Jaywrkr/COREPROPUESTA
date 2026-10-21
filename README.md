# COREPROPUESTA

Generador de propuestas de Coresolutions: convierte un configurador Lenovo (exportación de **Bid Platform** o cotización **DCSC**) en la propuesta oficial en Excel (hojas `Propuesta` y `_Lenovo`), con la misma plantilla, fórmulas, encabezado y firma de siempre.

Todo corre en el navegador: los archivos que se suben no salen del equipo del vendedor.

## Uso

1. Entrar con la clave del sistema.
2. Subir el configurador (.xlsx).
3. Revisar y editar la propuesta como en Excel (doble clic, Ctrl+C / Ctrl+V con Excel, clic derecho para insertar líneas).
4. Descargar: el Excel sale **cifrado con la misma clave** (cifrado estándar de Office, AES-256).

## Reglas de descripción (`reglas.json`)

Cuando un configurador trae un componente que la app no sabe describir, aparece en "Revisar antes de enviar". Para que salga siempre bien, agrega una regla en `reglas.json` (se edita directo en GitHub; Vercel publica solo al guardar):

```json
{
  "reglas": {
    "B8P9": { "texto": "{cant} puertos Ethernet 100Gb QSFP56,", "accion": "linea" },
    "C3K9": { "accion": "ignorar" }
  }
}
```

- La clave es el código de característica o P/N de Lenovo.
- `{cant}` se reemplaza por la cantidad por servidor.
- `"accion": "ignorar"` hace que el componente no se mencione.
- Una regla manda sobre las reglas incorporadas del motor.

## Aprendizaje: comentarios y correcciones

- **Comentar:** clic derecho en una línea → *Comentar esta línea…*. Se crea un issue en este repo (etiquetas `feedback` + `comentario`) con la línea, el comentario, el archivo y los códigos Lenovo que la generaron.
- **Correcciones automáticas:** al descargar, si el vendedor cambió líneas generadas, se crea un issue `correccion` con el antes → después.
- **Bandeja:** pestaña *Aprendizaje* (clave de administrador). Cada elemento se puede **convertir en regla** (commit a `reglas.json` + cierra el issue; Vercel publica en ~1 minuto) o **cerrar sin regla**.

### Configuración en Vercel (Settings → Environment Variables, entorno Production)

| Variable | Valor |
|---|---|
| `GITHUB_TOKEN` | Token *fine-grained* de GitHub, solo para este repo, con permisos **Issues: Read and write** y **Contents: Read and write** |
| `APP_CLAVE` | La clave de acceso de la app (la misma de los archivos) |
| `ADMIN_CLAVE` | Una clave distinta, solo para quien administra la bandeja |
| `GITHUB_REPO` | Opcional. Por defecto `Jaywrkr/COREPROPUESTA` |

Después de agregarlas, redeploy (Deployments → ⋯ → Redeploy). Sin `GITHUB_TOKEN` y `APP_CLAVE` la app funciona igual, solo sin comentarios ni bandeja.

### API

| Ruta | Uso |
|---|---|
| `GET /api/ping` | Indica si la bandeja está configurada |
| `POST /api/feedback` | Crea comentario/corrección (header `x-clave`) |
| `GET /api/feedback` | Lista pendientes (header `x-admin`) |
| `POST /api/regla` | Crea/edita/elimina una regla en `reglas.json` (header `x-admin`) |
| `POST /api/cerrar` | Cierra un pendiente sin regla (header `x-admin`) |

## Estructura

| Ruta | Qué es |
|---|---|
| `index.html` | La app ya armada (lo que publica Vercel) |
| `vault.json` | Plantilla oficial y firma, **cifradas** con la clave del sistema |
| `reglas.json` | Reglas de descripción del equipo |
| `api/` | Funciones de Vercel para comentarios, bandeja y reglas |
| `src/engine.js` | Motor: lee el configurador, arma la descripción y el Excel |
| `src/officecrypt.js` | Cifrado de Office (ECMA-376 Agile) para la descarga |
| `src/app.src.html` | Interfaz |
| `src/build.js` | Arma `index.html` |
| `src/make-vault.js` | Cifra plantilla y firma en `vault.json` |
| `tools/mktpl.py` | Prepara la plantilla base a partir de un archivo 2 real |

## Volver a armar

```bash
node src/build.js vault.json index.html
```

## Cambiar la clave, la firma o la plantilla

Requiere la plantilla en claro (no se guarda en el repo):

```bash
python3 tools/mktpl.py ARCHIVO2_REAL_DESCIFRADO.xlsx plantilla.xlsx
echo '{"firma":{"nombre":"...","cel":"...","email":"..."}}' > config.json
CLAVE='nueva-clave' node src/make-vault.js plantilla.xlsx config.json vault.json
node src/build.js vault.json index.html
```

No subas `plantilla.xlsx` ni `config.json` (están en `.gitignore`).
