# TechTeam · Assessment + Clasificación Técnica

Aplicación en **Streamlit** para analizar evaluaciones técnicas individuales, clasificar correctamente al técnico según la lógica del Excel corporativo y generar un **informe corporativo** con identidad visual TechTeam / Nutreco.

Esta versión está centrada en:

- **leer bien la calificación final** desde el propio Excel;
- mostrar resultados de assessment y comparativas;
- incorporar **branding corporativo**;
- proteger el acceso con **contraseña**;
- permitir descarga de **informe HTML corporativo**.

---

## Funcionalidades principales

### 1. Acceso restringido
La app solicita contraseña cada vez que se abre la sesión.

**Contraseña de esta versión:**

```text
TechTeam2026+
```

> Es una barrera básica de acceso en la interfaz.  
> Si en el futuro quieres un nivel de seguridad mayor, conviene moverla a `st.secrets`.

---

### 2. Lectura correcta de la calificación final
La calificación final **no se calcula con una media simple del score**.

La app se ciñe al criterio del Excel corporativo y toma la clasificación desde la hoja de base de datos / BBDD, usando la lógica que alimenta la celda final de ranking.

En esta versión, la clasificación se basa en la estructura del libro y en los cortes de clasificación de la hoja correspondiente, de forma que la app **debe coincidir con la calificación del Excel**.

Las categorías visibles son:

- **Básico**
- **Controla**
- **Supera**
- **Certificado**
- **Excelente**
- **Máster**

---

### 3. Assessment técnico
La app permite cargar archivos individuales en:

- `.xlsm`
- `.xlsx`

A partir del assessment:

- lee los indicadores técnicos;
- agrupa por áreas / troncos;
- calcula score global;
- compara con referencia;
- compara con benchmark / BBDD;
- detecta fortalezas;
- detecta debilidades.

---

### 4. Áreas del assessment
La lógica del assessment se estructura sobre cuatro áreas principales:

- **Alimentación**
- **Sanidad**
- **Manejo**
- **Herramientas**

---

### 5. Identidad corporativa
La app incorpora una cabecera visual corporativa con estos elementos:

- **Logo Nutreco**
- **Logo TechTeam**
- **Franja rosa corporativa**

La misma identidad visual se utiliza también en el informe HTML descargable.

---

### 6. Informe HTML corporativo
La app permite descargar un informe en **HTML** con plantilla corporativa.

Ese informe incluye:

- cabecera con logos;
- datos del técnico;
- clasificación final;
- resultados del assessment;
- interpretación general;
- fortalezas y debilidades;
- resumen técnico con formato visual de marca.

---

## Qué problema corrige esta versión

Esta versión corrige el error de las versiones previas en las que la clasificación visible podía salir de una lógica simplificada basada en el score medio.

Ahora la app debe alinearse con la lógica real del Excel corporativo y reflejar la **misma calificación final** que aparece en el libro fuente.

---

## Archivos necesarios en el repositorio

Debes dejar en la raíz del repositorio, como mínimo:

- `main.py`
- `requirements.txt`
- `README.md`

Y además, para mantener la identidad corporativa:

- `Logo Nutreco.jpg`
- `Logo TechTeam 2.jpg`
- `Solapa rosa.jpg`

Opcionalmente, puedes dejar assessments de ejemplo para pruebas.

---

## Estructura recomendada del repositorio

```text
/
├── main.py
├── requirements.txt
├── README.md
├── Logo Nutreco.jpg
├── Logo TechTeam 2.jpg
├── Solapa rosa.jpg
├── Technical Assessment Career Tool Ru Fco Marques.xlsm
└── ...
```

---

## Despliegue en Streamlit Community Cloud

### Paso 1. Subir a GitHub
Sube al repositorio:

- `main.py`
- `requirements.txt`
- `README.md`
- logos corporativos
- assessments de ejemplo opcionales

### Paso 2. Crear la app
En Streamlit Community Cloud:

1. conecta tu cuenta de GitHub;
2. pulsa **Create app**;
3. selecciona:
   - repositorio
   - rama
   - archivo principal: `main.py`
4. pulsa **Deploy**.

---

## Dependencias

El archivo debe llamarse exactamente:

```text
requirements.txt
```

Contenido recomendado:

```txt
streamlit>=1.44.0
pandas>=2.2.0
openpyxl>=3.1.0
plotly>=5.20.0
python-docx>=1.1.0
reportlab>=4.0.0
xlsxwriter>=3.2.0
```

> Si alguna versión concreta de la app usa exportación adicional o gráficos convertidos a imagen, puede requerir librerías extra.

---

## Cómo usar la app

1. Abre la app.
2. Introduce la contraseña.
3. Sube el archivo de assessment.
4. Revisa:
   - la clasificación final;
   - el score global;
   - las comparativas;
   - las fortalezas;
   - las debilidades.
5. Descarga el informe HTML corporativo.

---

## Qué debe comprobar el usuario

Conviene verificar especialmente:

- que la **calificación final de la app coincide con la del Excel**;
- que el archivo assessment mantiene la estructura esperada;
- que las fórmulas del Excel han sido recalculadas y guardadas antes de subirlo.

---

## Limitaciones conocidas

- La contraseña está integrada en el código, no en un sistema de secretos.
- La app depende de una estructura de Excel compatible con la plantilla corporativa.
- Si el Excel no se ha recalculado correctamente antes de guardarse, la clasificación puede no reflejar el estado esperado del libro.
- La exportación HTML depende de que los logos estén presentes en el repositorio.

---

## Recomendaciones

- Mantén los logos siempre en la raíz del proyecto.
- Usa assessments ya revisados y guardados en Excel.
- Revisa que la clasificación final coincida con la hoja de BBDD antes de distribuir el informe.
- Si quieres reforzar seguridad, mueve la contraseña a `st.secrets`.

---

## Objetivo de negocio

Esta app ayuda a:

- estandarizar la lectura del assessment técnico;
- asegurar que la clasificación final coincide con el criterio corporativo del Excel;
- presentar la información con imagen de marca;
- y compartir un informe claro, visual y coherente con la herramienta interna.

