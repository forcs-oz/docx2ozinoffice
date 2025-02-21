// <!> When you build jar with java 1.8
//import org.jetbrains.kotlin.gradle.dsl.JvmTarget
//import org.jetbrains.kotlin.gradle.tasks.KotlinCompile
//java {
//    sourceCompatibility = JavaVersion.VERSION_1_8
//    targetCompatibility = JavaVersion.VERSION_1_8
//}
//tasks.withType<KotlinCompile> {
//    compilerOptions {
//        jvmTarget.set(JvmTarget.JVM_1_8)
//    }
//}

plugins {
    kotlin("jvm") version "2.0.0"
}

group = "com.forcs.ozinoffice"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    testImplementation(kotlin("test"))
    implementation("com.documents4j:documents4j-local:1.1.13") // document4j
    implementation("com.documents4j:documents4j-transformer-msoffice-word:1.1.13")
    implementation("org.slf4j:slf4j-api:2.0.0")  // SLF4J API
    implementation("org.slf4j:slf4j-simple:2.0.0") // SLF4J Simple Logger Implementation
}

tasks.test {
    useJUnitPlatform()
}

tasks.jar {
    archiveBaseName.set("docx2ozinoffice")

    manifest {
        attributes["Main-Class"] = "MainKt"
    }

    val dependencies = configurations.runtimeClasspath.get().map { if (it.isDirectory) it else zipTree(it) }
    from(files(dependencies))

    duplicatesStrategy = DuplicatesStrategy.EXCLUDE
}