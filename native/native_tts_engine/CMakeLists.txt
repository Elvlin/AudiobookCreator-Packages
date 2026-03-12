cmake_minimum_required(VERSION 3.20)
project(native_tts_engine LANGUAGES C)

add_library(native_tts_engine SHARED native_tts_engine.c)
target_compile_definitions(native_tts_engine PRIVATE _CRT_SECURE_NO_WARNINGS)
set_target_properties(native_tts_engine PROPERTIES
    OUTPUT_NAME "native_tts_engine"
)
