# curso
Agrupacion de contenido

---

## Extractor de texto de PDFs tipo diapositivas

El script `extract_pdf_slides.py` extrae el texto de múltiples archivos PDF (típicamente presentaciones/diapositivas) y lo consolida en un único documento compacto.

### Formatos de salida

| Formato | Extensión | Notas |
|---------|-----------|-------|
| Word    | `.docx`   | Por defecto. Comprimido, fácil de editar. |
| Texto   | `.txt`    | El más pequeño. Sin formato. |

### Instalación de dependencias

```bash
pip install -r requirements.txt
```

### Uso

```bash
# Un directorio completo de PDFs → Word (por defecto)
python extract_pdf_slides.py mis_diapositivas/

# Archivos individuales → Word
python extract_pdf_slides.py tema1.pdf tema2.pdf -o resumen

# Un directorio → texto plano (el más compacto)
python extract_pdf_slides.py mis_diapositivas/ -o resumen -f txt

# Mezcla de archivos y directorios
python extract_pdf_slides.py intro.pdf avanzado/ -o curso_completo -f docx
```

### Opciones

| Opción | Descripción |
|--------|-------------|
| `input` | Uno o más archivos `.pdf` o directorios (se buscan PDFs de forma recursiva) |
| `-o`, `--output` | Nombre del archivo de salida sin extensión (por defecto: `contenido_extraido`) |
| `-f`, `--format` | Formato: `docx` (por defecto) o `txt` |
