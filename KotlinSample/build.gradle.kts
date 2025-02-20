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