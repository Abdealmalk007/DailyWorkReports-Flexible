plugins { id("com.android.application") }
android {
    namespace = "com.protech.dailyreports"
    compileSdk = 36
    defaultConfig {
        applicationId = "com.protech.dailyreports.flexible"
        minSdk = 24
        targetSdk = 35
        versionCode = 155
        versionName = "15.5.0-flexible"
    }
    buildTypes { release { isMinifyEnabled = false } }
    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }
}
