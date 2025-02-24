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
        /**
         * コンバーターの新しいインスタンスを生成します。
         * 컨버터의 새로운 인스턴스를 생성합니다.
         */
        fun get(): Docx2OZinOffice {
            return Docx2OZinOffice()
        }
    }

    private val logger: Logger = LoggerFactory.getLogger("[Docx2OZinOffice]")
    private var inputPaths: List<String>? = null
    private var formList: List<FormItem>? = null
    private var outputDir: File? = null

    /**
     * 指定ディレクトリ経路から置き換えの設定ファイル(JSON)と変換するドキュメントファイル(DOCX)を認識します。
     * 지정 디렉토리 경로로부터 치환설정파일(JSON)과 변환할 문서파일(DOCX)을 인식합니다.
     */
    fun from(inputPath: String): Docx2OZinOffice {
        inputPaths = null
        formList = null

        val inputDir = File(inputPath)
        val inputDirAbsolutePath = inputDir.absolutePath.replace("\\", "/")
        if (!inputDir.exists() || !inputDir.isDirectory) {
            logger.error("❌ No directory in the input path: $inputDirAbsolutePath")
            return this
        }

        /**
         * 指定ディレクトリ経路から置き換えの設定ファイル(JSON)を探して、置き換えの設定を読み込みます。
         * 지정 디렉토리 경로로부터 치환설정파일(JSON)을 찾고, 치환설정을 읽어들입니다.
         */

        val findResults = inputDir.list().filter{path -> path.endsWith(".json")};
        if (findResults.isEmpty()){
            logger.error("❌ No json file was found in the input directory")
            return this;
        }
        val jsonFile = File(inputDir.absolutePath + "/" + findResults.get(0))
        if (!jsonFile.exists() || !jsonFile.canRead()) {
            logger.error("❌ No json file was found in the input directory")
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
        val jsonFileAbsolutePath = jsonFile.absolutePath.replace("\\", "/")
        if (jsonList.isEmpty()) {
            logger.error("❌ Could not read the json file: $jsonFileAbsolutePath")
            return this
        }
        logger.info("✅ the json file was read successfully: $jsonFileAbsolutePath}")

        /**
         * 指定ディレクトリ経路から変換するドキュメントファイル(DOCX)をリストアップします。
         * 지정 디렉토리 경로로부터 변환할 문서파일(DOCX)을 리스트업합니다.
         */

        val docxPaths = inputDir.list().filter { it.endsWith(".docx") && !it.startsWith("~\$") }
        if (docxPaths.isEmpty()) {
            logger.error("❌ No docx file in the input path: $inputDirAbsolutePath")
            return this
        }
        logger.info("✅ ${docxPaths.size} docx file(s) found: $inputDirAbsolutePath")

        inputPaths = docxPaths.map { inputDirAbsolutePath + "/" + it }
        formList = jsonList

        return this
    }

    /**
     * 指定ディレクトリ経路に変換済みのドキュメントファイル(DOCX)を書く準備をします。
     * 지정 디렉토리 경로에 변환완료 문서파일(DOCX)을 쓸 준비를 합니다.
     */
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

    /**
     * 設定を初期化します。
     * 설정을 초기화합니다.
     */
    fun clear(): Docx2OZinOffice {
        inputPaths = null
        formList = null
        outputDir = null
        return this
    }

    /**
     * 設定どおりに変換作業を実行します。
     * 설정대로 변환작업을 실행합니다.
     */
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

    /**
     * 変換作業に用いられるビジュアルベーシックスクリプト(VBScript)を作成します。
     * 변환작업에 사용되는 비쥬얼베이직스크립트(VBScript)를 작성합니다.
     */
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

        /**
         * 読み込んだ置き換えの設定の内容をビジュアルベーシックスクリプト(VBScript)として記述します。
         * 읽어들인 치환설정의 내용을 비쥬얼베이직스크립트(VBScript)로 기술합니다.
         */

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

    /**
     * 一つのドキュメントファイル(DOCX)あたりの変換作業を実行します。
     * 한개의 문서파일(DOCX)에 대한 변환작업을 실행합니다.
     */
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