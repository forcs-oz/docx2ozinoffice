package com.forcs.ozinoffice

fun main(args: Array<String>) {
    val inputPath = args.getOrNull(0) ?: "src/test/resources/input"
    val outputPath = args.getOrNull(1) ?: "src/test/resources/output"

    Docx2OZinOffice.get()
        .from(inputPath)
        .to(outputPath)
        .run()
        .clear()
}