# Docx2OZinOffice
Converter sample for DOCX to OZ in Office

## Environment Settings (Windows 11)
1. MS Word
2. Kotlin IDE for Gradle Project
    - intelliJ IDEA
    - Oracle JDK 22
    - Gradle 8.8

## Usages

```
# Inputs:
<input-path>/*.docx
<input-path>/convert.json

# Outputs:
<output-path>/*.docx
```

1. Run main function

```
Usages: <input-path> <output-path>
```

2. Use API

```kt
import com.forcs.ozinoffice.Docx2OZinOffice

// ...

val inputPath: String = "<input-path>"
val outputPath: String = "<output-path>"
Docx2OZinOffice.get()
   .from(inputPath)
   .to(outputPath)
   .run()
   .clear()
```
