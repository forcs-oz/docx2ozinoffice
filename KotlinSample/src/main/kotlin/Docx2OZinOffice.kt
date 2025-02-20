package com.forcs.ozinoffice

import com.documents4j.api.DocumentType
import com.documents4j.api.IConverter
import com.documents4j.job.LocalConverter
import org.slf4j.LoggerFactory
import org.slf4j.Logger
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.nio.file.Files

class Docx2OZinOffice protected constructor() {
    companion object {
        fun get(): Docx2OZinOffice {
            return Docx2OZinOffice()
        }
    }

    private val logger: Logger = LoggerFactory.getLogger("[Docx2OZinOffice]")
    private var inputPaths: List<String>? = null
    private var jsonBytes: ByteArray? = null
    private var outputDir: File? = null

    fun from(inputPath: String): Docx2OZinOffice {
        this.inputPaths = null
        this.jsonBytes = null

        val inputDir = File(inputPath)
        if (!inputDir.exists() || !inputDir.isDirectory) {
            logger.error("❌ No directory in the input path: ${inputDir.absolutePath.replace("\\", "/")}")
            return this
        }

        val jsonFile = File(inputDir.absolutePath + "/convert.json")
        if (!jsonFile.exists() || !jsonFile.canRead()) {
            logger.error("❌ No convert.json file: ${jsonFile.absolutePath.replace("\\", "/")}")
            return this
        }

        var jsonBytes: ByteArray? = null
        try {
            FileInputStream(jsonFile).use { fis ->
                jsonBytes = fis.readAllBytes()
            }
        } catch (e: Throwable) {
            e.printStackTrace()
        }
        if (jsonBytes == null) {
            logger.error("❌ Could not read convert.json file: ${jsonFile.absolutePath.replace("\\", "/")}")
            return this
        }
        logger.info("✅ convert.json file was read successfully")

        val docxPaths = inputDir.list().filter { it.endsWith(".docx") }
        if (docxPaths.isEmpty()) {
            logger.error("❌ No docx file in the input path: ${inputDir.absolutePath.replace("\\", "/")}")
            return this
        }
        logger.info("✅ ${docxPaths.size} docx file(s) found: ${inputDir.absolutePath.replace("\\", "/")}")

        this.inputPaths = docxPaths.map { inputDir.absolutePath.replace("\\", "/") + "/" + it }
        this.jsonBytes = jsonBytes

        return this
    }

    fun to(outputPath: String): Docx2OZinOffice {
        this.outputDir = null

        try {
            val outputDir = File(outputPath)
            if (outputDir.exists()) {
                outputDir.deleteRecursively()
            }
            outputDir.mkdir()
            this.outputDir = outputDir;
        } catch (e: Throwable) {
            this.outputDir = null
        }

        return this
    }

    fun clear(): Docx2OZinOffice {
        inputPaths = null
        jsonBytes = null
        outputDir = null
        return this
    }

    fun run(): Docx2OZinOffice {
        if (inputPaths == null) {
            logger.error("❌ No input path yet.")
            return this
        }
        if (jsonBytes == null) {
            logger.error("❌ No convert.json yet.")
            return this
        }

        val tempDir = Files.createTempDirectory("docx2ozinoffice").toFile()
        logger.info("✅ Temporary Directory: ${tempDir.absolutePath.replace("\\", "/")}")

        var converter: IConverter? = null

        try {
            converter = LocalConverter.builder()
                .baseFolder(tempDir)
                .build()
            logger.info("✅ Converter was started successfully")

            if (this.prepareVBScript(tempDir)) {
                val count = inputPaths?.size
                inputPaths?.forEachIndexed() { index, it ->
                    val logPrefix = "[" + (index + 1) + "/$count] Convert: $it";
                    val outputPath = this.convert(converter, File(it))
                    if (outputPath.isEmpty()) {
                        logger.info("❌ $logPrefix")
                    } else {
                        logger.info("✅ $logPrefix  --->  $outputPath")
                    }
                }
            }
        } catch (e: Throwable) {
            e.printStackTrace()
        } finally {
            converter?.shutDown()
            tempDir.deleteRecursively();
        }
        logger.info("✅ Converter was terminated")

        return this
    }

    private fun prepareVBScript(tempDir: File): Boolean {
        val findResults = tempDir.list().filter{path -> path.endsWith(".vbs")};
        if (findResults.isEmpty()){
            logger.error("❌ VBScript was not found in the temporary directory")
            return false;
        }
        val vbsFile = File(tempDir.absolutePath + "/" + findResults.get(0))
        if (!vbsFile.exists() || !vbsFile.canRead() || !vbsFile.canWrite()) {
            logger.error("❌ VBScript was not found in the temporary directory")
            return false;
        }
        var vbsBytes: ByteArray? = null
        try {
            FileInputStream(vbsFile).use { fis ->
                vbsBytes = fis.readAllBytes()
            }
        } catch (e: Throwable) {
            e.printStackTrace()
        }
        if (vbsBytes == null) {
            logger.error("❌ Could not read VBScript")
            return false
        }
        logger.info("✅ VBScript was prepared")
        return true;
    }

    private fun convert(converter: IConverter, inputFile: File): String {
        if (inputPaths == null || jsonBytes == null || outputDir == null) {
            return ""
        }
        if (!inputFile.exists() || !inputFile.canRead()) {
            return ""
        }
        val outputPath = outputDir?.absolutePath?.replace("\\", "/") + "/" + inputFile.name;
        try {
            FileInputStream(inputFile).use { inputStream ->
                val outputFile = File(outputPath)
                FileOutputStream(outputFile).use { outputStream ->
                    converter
                        .convert(inputStream).`as`(DocumentType.DOCX)
                        .to(outputStream).`as`(DocumentType.DOCX)
                        .execute()
                }
            }
        } catch (e: Throwable) {
            e.printStackTrace()
            return ""
        }
        return outputPath;
    }
}