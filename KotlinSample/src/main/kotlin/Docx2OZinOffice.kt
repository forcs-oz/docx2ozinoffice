package com.forcs.ozinoffice

import com.documents4j.api.DocumentType
import com.documents4j.api.IConverter
import com.documents4j.job.LocalConverter
import com.google.gson.Gson
import com.google.gson.reflect.TypeToken
import org.slf4j.LoggerFactory
import org.slf4j.Logger
import java.io.*
import java.nio.file.Files

data class FormItem(val key: String, val formid: String)

class Docx2OZinOffice protected constructor() {
    companion object {
        fun get(): Docx2OZinOffice {
            return Docx2OZinOffice()
        }
    }

    private val logger: Logger = LoggerFactory.getLogger("[Docx2OZinOffice]")
    private var inputPaths: List<String>? = null
    private var formList: List<FormItem>? = null
    private var outputDir: File? = null

    fun from(inputPath: String): Docx2OZinOffice {
        inputPaths = null
        formList = null

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
        var jsonList = listOf<FormItem>()
        try {
            val jsonString = jsonFile.readText(Charsets.UTF_8)
            val listType = object : TypeToken<List<FormItem>>() {}.type
            jsonList = Gson().fromJson(jsonString, listType)
            jsonList = jsonList.filter { it.key.isNotBlank() }.distinctBy { it.key }
        } catch (e: Throwable) {
            e.printStackTrace()
        }
        if (jsonList.isEmpty()) {
            logger.error("❌ Could not read convert.json file: ${jsonFile.absolutePath.replace("\\", "/")}")
            return this
        }
        logger.info("✅ convert.json file was read successfully")

        val docxPaths = inputDir.list().filter { it.endsWith(".docx") && !it.startsWith("~\$") }
        if (docxPaths.isEmpty()) {
            logger.error("❌ No docx file in the input path: ${inputDir.absolutePath.replace("\\", "/")}")
            return this
        }
        logger.info("✅ ${docxPaths.size} docx file(s) found: ${inputDir.absolutePath.replace("\\", "/")}")

        inputPaths = docxPaths.map { inputDir.absolutePath.replace("\\", "/") + "/" + it }
        formList = jsonList

        return this
    }

    fun to(outputPath: String): Docx2OZinOffice {
        outputDir = null

        try {
            val outDir = File(outputPath)
            if (outDir.exists()) {
                outDir.deleteRecursively()
            }
            outDir.mkdir()
            outputDir = outDir;
        } catch (e: Throwable) {
            outputDir = null
        }

        return this
    }

    fun clear(): Docx2OZinOffice {
        inputPaths = null
        formList = null
        outputDir = null
        return this
    }

    fun run(): Docx2OZinOffice {
        if (inputPaths == null) {
            logger.error("❌ No input path yet.")
            return this
        }
        if (formList == null) {
            logger.error("❌ No convert.json yet.")
            return this
        }

        val tempDir = Files.createTempDirectory("docx2ozinoffice-").toFile()
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
        val vbsContentsFile = File("src/main/resources/word_convert.vbs");
        if (!vbsContentsFile.exists() || !vbsFile.canRead() || !vbsFile.canWrite()) {
            logger.error("❌ No VBScript contents")
            return false;
        }
        try {
            var vbsCont = vbsContentsFile.readText(Charsets.UTF_8)
            var addItemDictCode = ""
            if (formList != null) {
                formList!!.forEach { entry ->
                    addItemDictCode += "\n"
                    addItemDictCode += "\tSet itemDict = CreateObject(\"Scripting.Dictionary\")\n"
                    addItemDictCode += "\titemDict.Add \"key\", \"${entry.key}\"\n"
                    addItemDictCode += "\titemDict.Add \"formid\", \"${entry.formid}\"\n"
                    addItemDictCode += "\tdict.Add \"${entry.key}\", itemDict\n"
                }
                addItemDictCode += "\n"
            }
            vbsCont = vbsCont.replace("'@{#ADD_ITEMS#}@", addItemDictCode)
            vbsFile.writeText("\uFEFF" + vbsCont, Charsets.UTF_16LE) // UTF-16 LE BOM
        } catch (e: Throwable) {
            e.printStackTrace()
            logger.error("❌ Could not prepare VBScript")
            return false
        }

        logger.info("✅ VBScript was prepared")
        return true;
    }

    private fun convert(converter: IConverter, inputFile: File): String {
        if (inputPaths == null || formList == null || outputDir == null) {
            return ""
        }
        if (!inputFile.exists() || !inputFile.canRead()) {
            return ""
        }
        val jsonPath = inputFile.absolutePath.replace(".docx", ".json")
        val jsonFile = File(jsonPath)
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
            jsonFile.delete()
        } catch (e: Throwable) {
            e.printStackTrace()
            return ""
        }
        return outputPath;
    }
}