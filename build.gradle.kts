plugins {
    kotlin("jvm") version "1.6.20"
}

group = "com.cunningbird.sfedu.corp"
version = "1.0.0"

repositories {
    mavenCentral()
}

dependencies {
    implementation(kotlin("stdlib"))

    implementation("org.apache.poi:poi-ooxml:5.2.2")
}