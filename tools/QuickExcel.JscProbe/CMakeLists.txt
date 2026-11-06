# Development only. Never referenced by QuickExcel.slnx or ScriptWorker.targets.
# The JSCOnly Darwin framework layout does not add the framework header roots to
# JavaScriptCore's own include path, only the flattened private-header subdir.
if (APPLE)
    get_target_property(_qx_jsc_includes JavaScriptCore INCLUDE_DIRECTORIES)
    if (NOT _qx_jsc_includes)
        set(_qx_jsc_includes "")
    endif()
    list(APPEND _qx_jsc_includes "${JavaScriptCore_FRAMEWORK_HEADERS_DIR}" "${JavaScriptCore_PRIVATE_FRAMEWORK_HEADERS_DIR}")
    set_property(TARGET JavaScriptCore PROPERTY INCLUDE_DIRECTORIES "${_qx_jsc_includes}")
endif()
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
