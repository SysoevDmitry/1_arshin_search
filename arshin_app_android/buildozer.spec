[app]
title = ФГИС АРШИН
package.name = arshin_search
package.domain = com.arshin
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,csv
source.include_patterns = config/*.csv,*.png
version = 6.2
version.code = 1
requirements = python3,kivy,requests,pandas,openpyxl

orientation = portrait
fullscreen = 1
window.clearcolor = 1,1,1,1

android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 31
android.minapi = 26
android.ndk = 25b
android.sdk = 31
android.gradle_dependencies =

android.arch = arm64-v8a
android.allow_backup = True
android.logcat_filters = *:S python:D
android.presplash_color = #1B5E20

p4a.branch = develop
p4a.hostpython = python3

ios.kivy_ios_url = https://github.com/kivy/kivy-ios
ios.kivy_ios_branch = master

[buildozer]
log_level = 2
warn_on_root = 1
