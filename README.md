# Docx2OZinOffice
一般DOCXファイルをOZ-in-Office用ファイルへ転換する際に事前作業を処理するコンバーターのサンプルです。

일반 DOCX 파일을 OZ-in-Office용 파일로 전환할 시에 사전작업을 처리하는 컨버터 샘플입니다.

This is a converter sample for pre-processing DOCX to OZ-in-Office.


## 環境設定 / 환경설정 / Environment Settings

1. Windows 11
2. MS Word 2012
3. Kotlin IDE for Gradle Project
    - Oracle JDK 22
    - Gradle 8.8
    - intelliJ IDEA (optional, recommend)


## ビルド / 빌드 / Build

### ビルドの方法 / 빌드 방법 / How To Build

ターミナルで以下のコマンドを実行するか、IDEでGradleタスクの`clean`と`build`を実行します。

터미널에서 이하의 커맨드를 실행하거나、IDE에서 Gradle 태스크 `clean`과 `build`를 실행합니다.

Run the follow commands on the terminal, or the Gradle tasks `clean` and `build` on the IDE.

```shell
./gradlew clean build
```

### ビルドの結果 / 빌드 결과 / Build Results

ビルドが成功すると、`build/libs/docx2ozinoffice-{version}.jar`がエキスポートされます。

빌드가 성공하면, `build/libs/docx2ozinoffice-{version}.jar`가 익스포트됩니다.

If the build succeeds, `build/libs/docx2ozinoffice-{version}.jar` is exported.


### ビルドのカスタマイズ / 빌드 커스터마이징 / Build Customizing

`build.gradle.kts`を編集することで、ビルドされるJARファイルのバージョン名やファイル名、もしくはターゲットJDKバージョンの指定の変更が可能です。

`build.gradle.kts`를 편집함으로써, 빌드될 JAR 파일의 버전명 및 파일명의 변경, 혹은 타겟 JDK 버전의 지정 변경이 가능합니다.

By editing `build.gradle.kts`, you can change the version name, file name, or the target JDK version specification of the JAR file.

## 入力と結果 / 입력과 결과 / Inputs and Results

### 構成 / 구성 / Components

```
# Inputs:
<input-path>/setting.json
<input-path>/foo.docx
<input-path>/bar.docx

# Outputs:
<output-path>/foo.docx
<output-path>/bar.docx
```

テスト用リソースは`src/test/resources/input`で確認できます。

테스트용 리소스는 `src/test/resources/input`에서 확인할 수 있습니다.

The test resources are available at `src/test/resources/input`.

### 置き換えの設定 / 치환설정 / Replacement Settings

```javascript
[
   {
      "key": "@{Patient.ID}",    /* DOCXファイルに含まれる文字列の指定 */
      "formid": "patientid"      /* 指定の文字列が変換されたOZ-in-Officeコンポーネントの「formid」属性に与えられる値 */
   },
   {
      "key": "@{User.Addr}",     /* DOCX파일에 포함된 문자열의 지정 */
      "formid": "useraddr"       /* 지정된 문자열이 변환된 OZ-in-Office 컴포넌트의「formid」속성에 할당되는 값 */
   },
   {
      "key": "@{User.Name}",     /* Specifying the string included in the DOCX file */
      "formid": "username"       /* The value assigned to the 「formid」 property of the converted OZ-in-Office component with the specified string */
   },
   {
      "key": "@{患者_性別}",
      "formid": "patientsex"
   }
]
```

## 実行方法 / 실행방법 / How To Execute

### メイン関数を実行する / 메인 함수를 실행한다 / Run the main function

```
Usages: <input-path> <output-path>
```

### APIを呼び出す / API를 호출한다 / Call the API

#### Kotlin

```kt
import com.forcs.ozinoffice.Docx2OZinOffice

fun main() {
    val inputPath: String = "<input-path>"
    val outputPath: String = "<output-path>"
    Docx2OZinOffice.get()
        .from(inputPath)
        .to(outputPath)
        .run()
        .clear()
}
```

#### Java

```java
import com.forcs.ozinoffice.Docx2OZinOffice;

public class Main {
   public static void main() {
      String inputPath = "<input-path>";
      String outputPath = "<output-path>";
      
      Docx2OZinOffice.get()
             .from(inputPath)
             .to(outputPath)
             .run()
             .clear();
   }
}

```