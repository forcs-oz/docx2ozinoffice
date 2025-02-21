package com.forcs.ozinoffice

fun main(args: Array<String>) {
    if (args.size != 2) {
        println("====================================================")
        println("[Docx2OZinOffice] Usages: <input-dir> <output-dir>")
        println()
        println("Input Example:")
        println("\t<input-dir>/convert.json")
        println("\t<input-dir>/foo.docx")
        println("\t<input-dir>/bar.docx")
        println()
        println("Output Example:")
        println("\t<output-dir>/foo.docx")
        println("\t<output-dir>/bar.docx")
        return
    }

    val inputPath = args.getOrNull(0) ?: ""
    val outputPath = args.getOrNull(1) ?: ""

    Docx2OZinOffice.get()
        .from(inputPath)
        .to(outputPath)
        .run()
        .clear()
}