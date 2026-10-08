# Anexos V · Fajas UE (SENASA)

Aplicación web para generar la Declaración Jurada de fajas de seguridad para mercado UE
(Anexo V) a partir de las bajadas de scans. El equipo ingresa a un enlace, sube el Excel
y descarga los anexos, sin depender de nadie para generarlos.

## Qué hace

1. Recibe una o varias bajadas `.xlsx` (hoja `ag-grid`, columnas `Fecha`, `Faja1`, `Faja2`).
2. Genera un anexo por cada día con producción UE (series `FE` / `FF`; las `FC` se excluyen),
   con las mismas reglas y el mismo formato de la plantilla oficial.
3. Entrega: ZIP completo, PDF cronológico (una página por día) y cada Excel/PDF por separado.
4. Marca para confirmar los casos dudosos: fajas sueltas o muy alejadas del tramo principal,
   y volúmenes inusualmente altos o bajos. Los archivos idénticos se detectan y se usan una sola vez.

El pie «SERIE ___ Nº ___» y las firmas quedan en blanco para completar a mano.

## Estructura

Todos los archivos están en una sola carpeta, sin subcarpetas: se pueden subir juntos a GitHub.

```
app.py                              Aplicación Flask (rutas, acceso, descargas)
service.py                          Lectura de bajadas, controles, generación, PDF y ZIP
generador_anexos.py                 Reglas de negocio originales (day_tramos / build_day), sin cambios
plantilla_Fajas_UE_fixed.xlsx       Plantilla oficial con el logo SENASA corregido
base.html, index.html, resultado.html   Interfaz
Dockerfile, render.yaml, fonts.conf     Despliegue (LibreOffice, necesario para el PDF)
requirements.txt                    Dependencias de Python
```

## Publicar: GitHub y Render

**1. Subir a GitHub** (desde la carpeta del proyecto):

```bash
git init
git add .
git commit -m "App Anexos V Fajas UE"
git branch -M main
git remote add origin https://github.com/<usuario>/<repositorio>.git
git push -u origin main
```

Se recomienda un repositorio **privado**.

**2. Desplegar en Render**

1. En Render: **New → Web Service** y conectar el repositorio de GitHub.
2. Render detecta el `Dockerfile` (o usar **New → Blueprint** para tomar `render.yaml`).
3. En **Environment**, cargar `APP_PASSWORD` con la contraseña de acceso del equipo.
   El usuario por defecto es `fajas` (se cambia con `APP_USER`).
4. **Create Web Service**. La primera construcción tarda varios minutos (instala LibreOffice).
5. Compartir con el equipo el enlace `https://<nombre>.onrender.com`.

Cada `git push` a `main` vuelve a desplegar automáticamente.

## Variables de entorno

| Variable | Por defecto | Uso |
|---|---|---|
| `APP_PASSWORD` | *(vacía)* | Si se define, se pide usuario y contraseña. **Definirla**: el enlace es público. |
| `APP_USER` | `fajas` | Usuario de acceso. |
| `VOLUMEN_ALTO` | `4000` | Fajas/día por encima de las cuales se pide confirmar la bajada. |
| `VOLUMEN_BAJO` | `600` | Fajas/día por debajo de las cuales se pide confirmar la bajada. |
| `TRAMO_CHICO` | `5` | Un tramo secundario con hasta esta cantidad de fajas se marca como posible escaneo aislado. |
| `DISTANCIA_LEJANA` | `10000` | Un tramo secundario más lejos que esto del principal se marca como posible escaneo antiguo. |
| `MAX_DIAS` | `62` | Máximo de días por carga. |
| `MAX_UPLOAD_MB` | `60` | Tamaño máximo de la carga. |
| `JOB_TTL_HORAS` | `6` | Horas que se conservan los resultados antes de borrarse. |

## Si aparece «Internal Server Error»

1. Abrir `https://<nombre>.onrender.com/diagnostico`. Indica si falta alguna carpeta o archivo
   (por ejemplo `index.html` si la subida a GitHub quedó incompleta) y si LibreOffice está instalado.
   Agregando `?prueba_pdf=1` prueba además la conversión a PDF y su duración.
2. En Render, entrar al servicio y abrir **Logs**: cada error queda registrado con un código y el detalle completo
   (líneas que empiezan con `Traceback`).
3. Verificar en GitHub que estén **todos** los archivos del listado de arriba (son 12; `.gitignore` y `.dockerignore` son opcionales). Atención a `base.html`, `index.html`, `resultado.html` y a la plantilla `.xlsx`.

## Consideraciones del plan gratuito de Render

- El servicio se suspende tras un período sin uso: el primer acceso puede demorar cerca de un minuto.
- Tiene 512 MB de memoria. La conversión a PDF con LibreOffice es lo más pesado y se ejecuta de a una
  por vez. Si el servicio se reinicia por falta de memoria al cargar muchos días, pasar al plan **Starter**
  o cargar menos días por vez. Los Excel se generan igual aunque falle el PDF.
- El almacenamiento es temporal: las bajadas subidas no se conservan y los resultados se eliminan solos.

## Ejecutar en una computadora (opcional)

```bash
pip install -r requirements.txt
python app.py          # http://localhost:8000
```

Para el PDF se necesita LibreOffice instalado (`soffice` en el PATH). Sin él, la app entrega solo los Excel.

## Mantenimiento

- Cambiar una regla de cálculo: editar `generador_anexos.py` y hacer `git push`.
- Cambiar el formato: reemplazar `plantilla_Fajas_UE_fixed.xlsx`.
- Ajustar la sensibilidad de los controles: modificar las variables de entorno en Render.
