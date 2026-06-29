# pbImageOCR — OCR de imágenes y PDFs con PowerBuilder 🔎

![PowerBuilder](https://img.shields.io/badge/PowerBuilder-2025-2D6CDF?style=flat-square)
![.NET](https://img.shields.io/badge/.NET-10-512BD4?style=flat-square&logo=dotnet&logoColor=white)
![Tesseract](https://img.shields.io/badge/OCR-Tesseract-5C3EE8?style=flat-square)
![PDFtoImage](https://img.shields.io/badge/PDF-PDFtoImage-1f6feb?style=flat-square)
![Blog](https://img.shields.io/badge/blog-rsrsystem-FF5722?style=flat-square&logo=blogger&logoColor=white)

## 📋 ¿Qué es esto?

Un ejemplo PowerBuilder para hacer **OCR** (reconocer texto dentro de una imagen) usando un
**capturador con el cursor**: seleccionáis con el ratón un trozo de la pantalla, lo recortáis, y
PowerBuilder os devuelve el texto que había ahí dentro. También podéis partir de un **PDF** o de
una imagen que tengáis en el **portapapeles**.

¿Y cómo lo logra PowerBuilder, que de OCR no tiene ni idea? Pues **apoyándose en .NET**: cargamos
**tres librerías** .NET como `dotnetobject` con el **.NET DLL Importer** de PB, y cada una se usa
desde PowerScript como un objeto nativo más. Cada una hace su parte del flujo:

| Librería .NET | Qué hace | Métodos clave |
|---------------|----------|---------------|
| **`ImageOCR`** | El OCR de verdad: imagen → texto (motor **Tesseract**) | `ConvertImageToString(img)`, `ConvertImageToTxt(img, txt)` |
| **`ImageFromPdf`** | Rasteriza un PDF a imagen (paso previo para hacerle OCR) | `PdfToPng(pdf)`, `PdfToBmp(pdf)` |
| **`ImageFromClipboard`** | Vuelca a fichero la imagen que haya en el **portapapeles** | `GetClipboardImage()` |

La gracia didáctica es esa: un objetivo (sacar texto de cualquier cosa) repartido en tres piezas
.NET pequeñas y reutilizables, que PowerBuilder orquesta. Todas exponen además un `GetLastError()`
para que PowerBuilder reciba el fallo como texto en vez de pelearse con una excepción .NET.

## 🔗 Motor .NET

El trabajo lo hacen las **tres librerías** .NET de la solución **`ImageOCR`**:

- Se **despliegan** ya compiladas en las carpetas `DotNet\ImageOCR\`, `DotNet\ImageFromPdf\` y
  `DotNet\ImageFromClipboard\` de este propio ejemplo, para que clones, compiles y funcione sin
  tocar nada más.
- Se **consumen** desde PowerBuilder como `dotnetobject` (proxies creados con el .NET DLL Importer).
- El **código fuente** de las tres vive en `Blog\Net10\ImageOCR` (antes estaba en `Net8`,
  subproyectos `ImageOCR`, `ImageFromPdf` e `ImageFromClipboard`) y se recompila/despliega con el
  script **`desplegar_dotnet.bat`** (hace `dotnet publish` y espeja las DLLs a la carpeta `DotNet`
  de cada ejemplo).
- Repo del proyecto .NET (Visual Studio 2022): <https://github.com/rasanfe/ImageOCR>

> 🆕 **Nota didáctica (migración a .NET 10):** `ImageFromPdf` ya **no usa el abandonado PdfiumViewer
> (2018)**. Ahora rasteriza con **[PDFtoImage](https://www.nuget.org/packages/PDFtoImage)** (MIT,
> mantenido, PDFium + SkiaSharp). El OCR sigue con **Tesseract** (idioma `spa`); recordad llevar los
> datos de idioma (`tessdata`, p. ej. `spa.traineddata`) junto a la DLL. Si recompiláis, **volved a
> desplegar**.

## 🛠️ Requisitos

- **PowerBuilder 2025** para abrir y compilar la solución.
- **.NET 10 Runtime** instalado en la máquina → <https://dotnet.microsoft.com/en-us/download/dotnet/10.0>
- Las carpetas `DotNet\ImageOCR\`, `DotNet\ImageFromPdf\` y `DotNet\ImageFromClipboard\` con las
  DLLs desplegadas (ya vienen en el repo), incluyendo `tessdata` con el idioma para el OCR.

## ▶️ Cómo probarlo

1. Clona el repo y abre `pbImageOCR.pbsln` con PowerBuilder 2025.
2. Compila (Full Build) y ejecuta.
3. Usa el **capturador**: arrastra el cursor para recortar una zona de pantalla con texto y mira
   cómo te devuelve el contenido reconocido.
4. Prueba también partiendo de un **PDF** (se rasteriza y se le hace OCR) o pegando una imagen
   desde el **portapapeles**.

## 🔗 Repo PowerBuilder

<https://github.com/rasanfe/pbImageOCR>

## 🙌 Créditos

- Gracias a **Oscar Francisco Hernández V.** por sus ideas para crear el capturador.
- Gracias a **Topwiz Software** por sus ejemplos:
  - Bitmap: <https://www.topwizprogramming.com/freecode_bitmap.html>
  - Resize Response: <https://www.topwizprogramming.com/freecode_resize_response.html>

---

> ¡Nos vemos en el próximo artículo! Y recuerda: en PowerBuilder, los límites solo están en nuestra imaginación. 🚀

📨 **Blog:** <https://rsrsystem.blogspot.com/>
