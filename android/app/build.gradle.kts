plugins {
  alias(libs.plugins.android.application)
}

android {
  namespace = "com.jicorporate.gfmp"
  compileSdk = 34
  defaultConfig {
    applicationId = "com.jicorporate.gfmp"
    minSdk = 24
    targetSdk = 34
    versionCode = 7
    versionName = "1.6"
  }

  buildTypes {
    release {
      isMinifyEnabled = false
      proguardFiles(getDefaultProguardFile("proguard-android-optimize.txt"), "proguard-rules.pro")
    }
  }
  compileOptions {
    sourceCompatibility = JavaVersion.VERSION_17
    targetCompatibility = JavaVersion.VERSION_17
  }
  buildFeatures {
    compose = false
    aidl = false
    buildConfig = false
    shaders = false
  }

  packaging {
    resources {
      excludes += "/META-INF/{AL2.0,LGPL2.1}"
    }
  }
}

kotlin {
  jvmToolchain(17)
}

dependencies {
  implementation("androidx.core:core-ktx:1.13.1")
  implementation("androidx.activity:activity:1.9.1")
  implementation("androidx.webkit:webkit:1.11.0")
}
