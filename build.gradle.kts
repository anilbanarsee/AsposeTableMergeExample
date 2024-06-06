plugins {
    id("java")
}

group = "org.example"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
    maven("https://releases.aspose.com/java/repo/")
}

dependencies {
    testImplementation(platform("org.junit:junit-bom:5.10.0"))
    testImplementation("org.junit.jupiter:junit-jupiter")
    implementation("com.aspose", "aspose-words", "22.11", classifier = "jdk17")
}

tasks.test {
    useJUnitPlatform()
}