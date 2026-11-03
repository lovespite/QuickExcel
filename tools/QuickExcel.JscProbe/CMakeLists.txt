# Development only. Never referenced by QuickExcel.slnx or ScriptWorker.targets.
target_sources(JavaScriptCore PRIVATE "${QX_JSC_PROBE_SOURCE}/../../QuickExcel.Scripting/native/jsc/bridge.cpp")
add_executable(quickexcel-jsc-probe "${QX_JSC_PROBE_SOURCE}/probe.cpp")
target_compile_features(quickexcel-jsc-probe PRIVATE cxx_std_20)
target_include_directories(quickexcel-jsc-probe PRIVATE
    $<TARGET_PROPERTY:JavaScriptCore,INCLUDE_DIRECTORIES>
    "${JavaScriptCore_FRAMEWORK_HEADERS_DIR}"
    "${JavaScriptCore_PRIVATE_FRAMEWORK_HEADERS_DIR}")
target_link_libraries(quickexcel-jsc-probe PRIVATE JavaScriptCore)
if(WIN32)
    target_link_libraries(quickexcel-jsc-probe PRIVATE psapi)
endif()
