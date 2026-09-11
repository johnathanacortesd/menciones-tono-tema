# Menciones: Tono + Tema + Sub-tema

App de Streamlit que analiza un export de monitoreo de medios (GlobalNews u otro) y devuelve el
mismo XLSX con tres columnas nuevas por mención:

| Columna | Qué es |
|---|---|
| **Tono** | `Positivo` / `Neutro` / `Negativo`, medido sobre la entidad y/o su vocero (marca + alias) |
| **Tema** | Cubo amplio que agrupa sub-temas parecidos (lista cerrada por cliente, **sin "Otros"**) |
| **Sub-tema** | Frase nominal de 3 a 7 palabras que resume el hecho |

Subes el archivo, escribes la marca, los voceros y los alias, y descargas el XLSX con las tres
columnas y cuatro hojas: `Menciones`, `Temas (agrupa Sub-temas)`, `Grupos` y `Resumen`.

## Por qué acierta (y por qué un prompt suelto no)

La precisión no viene del prompt: viene de cinco piezas deterministas alrededor del modelo.

1. **Agrupación antes de etiquetar.** Las notas iguales o casi iguales (el mismo comunicado en 20
   medios) se agrupan y se etiquetan **una sola vez**. Por eso dos notas idénticas no pueden salir con
   tonos distintos. Es el error más común al etiquetar fila por fila con una API.
2. **Primero el sub-tema, después el tono.** El modelo resume el hecho y luego lo juzga. Además, los
   sub-temas ya asignados se pasan en cada lote como `CANDIDATOS`, así los hechos repetidos reutilizan
   **exactamente** el mismo texto y no aparecen variantes del mismo hecho.
3. **Validador con reglas duras + reparación.** Cada etiqueta se revisa (3-7 palabras, sin verbo
   conjugado al inicio, sin terminar en preposición, sin `:` `;` `|` ni comillas, sin rótulos vacíos
   como "noticias generales", tono dentro del vocabulario) y los fallos vuelven al modelo para que los
   corrija. Sin este paso quedan etiquetas de 9 palabras y titulares recortados.
4. **El Tema no lo inventa el modelo.** Los Temas son una lista cerrada de cubos por tipo de cliente y
   se asignan por reglas léxicas, mirando primero el sub-tema y solo después el título (y solo si el
   título mide ≤160 caracteres, porque hay exportes donde esa columna trae el cuerpo completo y
   palabras incidentales como *contratos* meten notas en el cubo equivocado).
5. **"Otros" no se entrega.** Si ningún cubo sirve, el modelo elige dentro de la lista o propone un
   cubo nuevo **específico**; la app rechaza `Otros`, `Información general`, `Actividad institucional`
   y cualquier rótulo vacío, y bloquea la descarga mientras quede un grupo sin cubo.

## Uso local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre `http://localhost:8501`.

## Despliegue en Streamlit Cloud

1. Sube el repo a GitHub.
2. En [share.streamlit.io](https://share.streamlit.io) elige el repo y `app.py`.
3. (Opcional) Guarda la API key como secreto en vez de escribirla cada vez:
   `Settings → Secrets` con

   ```toml
   [general]
   llm_api_key = "tu-key"
   proveedor = "Groq"
   modelo = "llama-3.3-70b-versatile"
   ```

   Si el archivo `secrets.toml` no existe, la app funciona igual: la key se escribe en la barra
   lateral y vive solo en la sesión del navegador.

## Configuración en la app

- **Cliente**: entidad, voceros y alias. Los alias importan: si falta la forma en que los medios
  nombran a la entidad (sigla, cargo, "el gremio avicultor", "la avicultura colombiana"), el tono sale
  Neutro de más o de menos.
- **Criterio del tono**: `Aspectual estricto` (Negativo solo con crítica dirigida a la entidad; ante
  duda, Neutro) o `Favorabilidad del sector` (para gremios: cuenta cómo queda parado el sector aunque
  la entidad no sea el actor).
- **Columnas**: la app detecta título, texto e id por nombre (Título/Titular, CuerpoEs/Contenido/Texto,
  NoticiaId/Id). Se pueden cambiar.
- **Agrupación**: el umbral de similitud de titulares (por defecto 92 %) y de cuerpos (85 %) son los
  controles que deciden cuántas notas se consideran "la misma". Bajarlos fusiona campañas publicadas
  por muchos medios; subirlos separa notas parecidas pero distintas.
- **Temas**: dos listas listas para usar (`Gobierno territorial`, 21 cubos; `Gremio o sector`, 16
  cubos) y un editor JSON para adaptarlas. El orden de las reglas es la prioridad: lo específico
  antes que lo genérico.

## Salida

- `Menciones`: todas las columnas del export + `Fila original`, `Grupo de similitud`, `Tono`, `Tema`,
  `Sub-tema` y `Criterio del tono`.
- `Temas (agrupa Sub-temas)`: cada cubo con sus grupos, menciones y sub-temas.
- `Grupos`: el grupo con su título representativo y su etiqueta (para auditar).
- `Resumen`: conteos por tono y por cubo, largo medio del sub-tema, uniformidad y control de calidad
  (los avisos de reparación del validador).

## Límites honestos

- El tono es un juicio. Las reglas cubren la mayoría de los casos, pero los casos límite (una crítica
  irónica, una obra anunciada y nunca ejecutada) pueden diferir de una lectura humana. La app deja
  reasignar antes de descargar.
- El modelo no es determinista al 100 % aunque la temperatura sea 0: guarda el XLSX que entregas como
  versión final del período.
- El costo depende del modelo y del número de grupos: 750 menciones suelen quedar en ~570 grupos, y con
  15 grupos por llamada son ~38 llamadas + reparaciones. Con un modelo pequeño (gpt-4.1-mini,
  llama-3.3-70b) el costo es de centavos.
- La agrupación por cuerpo exige al menos 30 cinco-gramas: en notas muy cortas (radio/TV sin cuerpo)
  solo agrupa por titular. Es intencional, para no fusionar notas distintas.

## Autotest (sin API key)

```bash
python autotest.py
```

Corre el pipeline completo con un modelo simulado que devuelve etiquetas malas a propósito y verifica:
agrupación, validador + reparación, prohibición de "Otros", rechazo de rótulos vacíos como cubo nuevo,
y el XLSX de salida (hojas, uniformidad por grupo, largo de los sub-temas).

## Estructura

```
app.py            todo la app: utilidades, agrupación, validador, taxonomías, capa LLM, XLSX, UI
autotest.py       verificación sin API key
requirements.txt
```

## Licencia

MIT.
