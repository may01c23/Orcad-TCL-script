# OrCAD Capture 格式刷脚本 (最终增强版)
# Version: 0.1
# ===================================================================
#  初始化环境
# ===================================================================
package require Tk
wm withdraw .

# ===================================================================
#  全局变量
# ===================================================================
array set g_copied_format { type -1 }
set g_copied_display_props [list]
array set g_apply_flags { color 1 rotation 1 mirror 1 display_type 1 layout 1 font 1}

# 用于确保类型名称映射只初始化一次的全局变量
set g_type_names_initialized 0
array set g_type_names {}

# ===================================================================
#  函数定义
# ===================================================================
# 从键值列表中获取指定键值的值
proc get_from_kv_list {kv_list key} {
    set idx [lsearch -exact $kv_list $key]; if {$idx == -1} { return "" }; return [lindex $kv_list [expr {$idx + 1}]]
}

# 初始化对象类型名称映射 (使用硬编码ID)
proc initialize_type_names {} {
    global g_type_names_initialized g_type_names
    if {$g_type_names_initialized} { return }

    # 使用硬编码的ID，定义脚本支持的对象类型
    array set g_type_names [list \
        12 "Block"\
        13 "Part" \
        23 "Port" \
        37 "Power" \
        38 "Off Page" \
        39 "Part Properties" \
        49 "Net Alias" \
    ]
    set g_type_names_initialized 1
}

# 将对象类型ID转换为可读字符串
proc get_object_type_string { type_id } {
    global g_type_names
    initialize_type_names

    if {[info exists g_type_names($type_id)]} {
        return $g_type_names($type_id)
    } else {
        # 对于列表之外的类型，返回一个通用提示
        puts "\r\n type id: $type_id. \r\n"
        return
    }
}

# 获取显示属性，适用于Part等多属性对象
proc get_all_display_props { obj } {
    set props_list [list]; set lStatus [DboState]
    if {[catch {$obj NewDisplayPropsIter $lStatus} props_iter]} { return "" }
    set prop [$props_iter NextProp $lStatus]
    while {$prop != "NULL"} {
        set prop_info [list]; set prop_name_cstr [DboTclHelper_sMakeCString]; $prop GetName $prop_name_cstr
        lappend prop_info "name" [DboTclHelper_sGetConstCharPtr $prop_name_cstr]
        lappend prop_info "color" [$prop GetColor $lStatus]
        set lFont [DboTclHelper_sMakeLOGFONT]; $prop GetFont $::DboLib_DEFAULT_FONT_PROPERTY $lFont; lappend prop_info "font" $lFont
        lappend prop_info "rotation" [$prop GetRotation $lStatus]
        lappend prop_info "display_type" [$prop GetDisplayType $lStatus]
        lappend prop_info "location" [$prop GetLocation $lStatus]
        lappend props_list $prop_info; set prop [$props_iter NextProp $lStatus]
    }
    delete_DboDisplayPropsIter $props_iter
    
    return $props_list
}

# 检查对象是否支持某个特定的方法
proc object_has_capability {obj capability} {
    set lStatus [DboState]; return [expr {![catch {$obj $capability $lStatus} _]} ]
}

#  同步更新选项的后台变量和界面状态
proc update_option_state {option_list is_enabled} {
    global g_apply_flags
    
    # 遍历列表中的每一个选项
    foreach option_name $option_list {
        set g_apply_flags($option_name) $is_enabled

        set state [expr {$is_enabled ? "normal" : "disabled"}]
        .format_painter.main_frame.options.cb_$option_name configure -state $state
    }
}

#  复制格式
proc copy_format {} {
    global g_copied_format g_copied_display_props g_type_names
    if {![winfo exists .format_painter]} { return }
    set selected_objects [GetSelectedObjects]; if {[llength $selected_objects] != 1} {tk_messageBox -icon error -title "错误" -message "请只选择一个源对象！"; return}
    
    set source_obj [lindex $selected_objects 0]; set lStatus [DboState]
    set current_type_id [$source_obj GetObjectType]

    initialize_type_names
    if {![info exists g_type_names($current_type_id)]} {
        set unsupported_type_name [get_object_type_string $current_type_id]
        tk_messageBox -icon error -title "操作不支持" -message "您选择的对象类型不支持格式刷操作。"
        # 确保不保留上一次的复制信息
        array unset g_copied_format; set g_copied_format(type) -1
        .format_painter.main_frame.feedback_label configure -text "源对象: 无"
        return
    }

    array unset g_copied_format; set g_copied_display_props [list]
    set g_copied_format(type) $current_type_id
    
    set feedback_items [list]
    
    #  优先获取位号，如果失败则获取通用名称
    set obj_display_name ""; set name_cstr [DboTclHelper_sMakeCString]
    if {![catch {$source_obj GetReferenceDesignator $name_cstr}]} { set obj_display_name [DboTclHelper_sGetConstCharPtr $name_cstr] }
    if {$obj_display_name == "" && ![catch {$source_obj GetName $name_cstr}]} { set obj_display_name [DboTclHelper_sGetConstCharPtr $name_cstr] }
    if {$obj_display_name == ""} { set obj_display_name "Selected Object" }

    # 根据对象类型执行不同的复制逻辑
    if {$g_copied_format(type) == 49 || $g_copied_format(type) == 39} { 
        # 处理纯文本对象
        set g_copied_format(is_text) 1
        set prop_info [list]

        if {$g_copied_format(type) == 49 } {
            # 禁用不适用于net alias的选项
            update_option_state {display_type layout} 0

            lappend prop_info "font"  [$source_obj GetFont $lStatus]
            lappend prop_info "color" [$source_obj GetColor $lStatus]

            lappend feedback_items "字体" "颜色"
            update_option_state {font color} 1
        } else {
            # 禁用不适用于Part Properties的选项
            update_option_state "layout" 0

            set lFont [DboTclHelper_sMakeLOGFONT]; $source_obj GetFont $::DboLib_DEFAULT_FONT_PROPERTY $lFont; lappend prop_info "font" $lFont
            lappend prop_info "color" [$source_obj GetColor $lStatus]
            lappend prop_info "display_type" [$source_obj GetDisplayType $lStatus]

            lappend feedback_items "字体" "颜色" "可见性"
            update_option_state {font color display_type rotation} 1
        }

        lappend g_copied_display_props $prop_info
    } else {
        set g_copied_format(is_text) 0
        
        if {[object_has_capability $source_obj "NewDisplayPropsIter"]} {
            set g_copied_display_props [get_all_display_props $source_obj]
            if {[llength $g_copied_display_props] > 0} { 
                lappend feedback_items "字体" "颜色" "可见性" "属性布局"
                update_option_state {font color display_type layout} 1
            }
        } else { update_option_state {font color display_type layout} 0 }
    }
    if {[object_has_capability $source_obj "GetRotation"]} {
        set g_copied_format(rotation) [$source_obj GetRotation $lStatus]; lappend feedback_items "旋转"; update_option_state "rotation" 1
    } else { update_option_state "rotation" 0 } 

    if {[object_has_capability $source_obj "GetMirror"]} {
        set g_copied_format(mirror) [$source_obj GetMirror $lStatus]; lappend feedback_items "镜像"; update_option_state "mirror" 1
    } else { update_option_state "mirror" 0 }

    puts $g_copied_display_props

    set feedback_str "已复制: [join [lsort -unique $feedback_items] ", "]"
    .format_painter.main_frame.feedback_label configure -text "源对象: $obj_display_name ([get_object_type_string $g_copied_format(type)])\n$feedback_str"
}

#  粘贴格式
proc paste_format {} {
    global g_copied_format g_copied_display_props g_apply_flags
    if {$g_copied_format(type) == -1} {tk_messageBox -icon warning -title "警告" -message "请先复制一个格式。"; return}
    set selected_objects [GetSelectedObjects]; if {[llength $selected_objects] == 0} {tk_messageBox -icon error -title "错误" -message "请至少选择一个目标对象。"; return}

    # 粘贴前，首先验证所有选定对象的类型是否匹配
    set source_type $g_copied_format(type)
    foreach target_obj $selected_objects {
        if {[$target_obj GetObjectType] != $source_type} {
            set source_type_name [get_object_type_string $source_type]
            set target_type_name [get_object_type_string [$target_obj GetObjectType]]
            
            set obj_display_name ""; set name_cstr [DboTclHelper_sMakeCString]
            if {![catch {$target_obj GetReferenceDesignator $name_cstr}]} { set obj_display_name [DboTclHelper_sGetConstCharPtr $name_cstr] }
            if {$obj_display_name == "" && ![catch {$target_obj GetName $name_cstr}]} { set obj_display_name [DboTclHelper_sGetConstCharPtr $name_cstr] }
            if {$obj_display_name != ""} { set obj_display_name " '$obj_display_name'"}
            
            set error_msg "类型不匹配！\n\n此格式刷适用于 “${source_type_name}” 类型的对象。\n您选择的对象${obj_display_name} (类型: ${target_type_name}) 与之不兼容。"
            tk_messageBox -icon error -title "格式粘贴错误" -message $error_msg
            return ; # 发现不匹配的对象，立即中断整个操作
        }
    }

    set lStatus [DboState]; set DO_NOT_DISPLAY 0
    set applied_count 0

    puts $g_copied_display_props
    
    foreach target_obj $selected_objects {
        incr applied_count

        if {[info exists g_copied_format(is_text)] && $g_copied_format(is_text)} {
            # 专门用于粘贴纯文本格式的逻辑
            set src_prop_info [lindex $g_copied_display_props 0]
            if {$g_apply_flags(font)}   { if {![catch {$target_obj SetFont [get_from_kv_list $src_prop_info "font"]}]} {} }
            if {$g_apply_flags(color)}  { if {![catch {$target_obj SetColor [get_from_kv_list $src_prop_info "color"]}]} {} }
            if {$g_apply_flags(display_type)} {
                puts 123
                set source_display_type [get_from_kv_list $src_prop_info "display_type"]
                if {[info exists source_display_type] && $source_display_type ne ""} {
                    if {![catch {$target_obj SetDisplayType $source_display_type}]} {}
                }
            }
        } else {
            if {[object_has_capability $target_obj "NewDisplayPropsIter"]} {
                if {$g_apply_flags(display_type)} {
                    set source_visible_prop_names [list]
                    foreach src_prop_info $g_copied_display_props { if {[get_from_kv_list $src_prop_info "display_type"] != $DO_NOT_DISPLAY} { lappend source_visible_prop_names [get_from_kv_list $src_prop_info "name"] } }
                    set target_props_iter [$target_obj NewDisplayPropsIter $lStatus]; set prop_to_check [$target_props_iter NextProp $lStatus]
                    while {$prop_to_check != "NULL"} {
                        set prop_name_cstr [DboTclHelper_sMakeCString]; $prop_to_check GetName $prop_name_cstr; set prop_name [DboTclHelper_sGetConstCharPtr $prop_name_cstr]
                        if {[$prop_to_check GetDisplayType $lStatus] != $DO_NOT_DISPLAY && [lsearch -exact $source_visible_prop_names $prop_name] == -1} { $prop_to_check SetDisplayType $DO_NOT_DISPLAY }
                        set prop_to_check [$target_props_iter NextProp $lStatus]
                    }
                }
                if {[llength $g_copied_display_props] > 0} {
                    foreach src_prop_info $g_copied_display_props {
                        set prop_name [get_from_kv_list $src_prop_info "name"]; set prop_name_cstr [DboTclHelper_sMakeCString $prop_name]; set target_prop [$target_obj GetDisplayProp $prop_name_cstr $lStatus]
                        if {$target_prop == "NULL"} { if {$g_apply_flags(color) || $g_apply_flags(rotation) || $g_apply_flags(display_type) || $g_apply_flags(font) || $g_apply_flags(layout)} { set font_to_apply [get_from_kv_list $src_prop_info "font"]; set location [get_from_kv_list $src_prop_info "location"]; set rotation [get_from_kv_list $src_prop_info "rotation"]; set color [get_from_kv_list $src_prop_info "color"]; set target_prop [$target_obj NewDisplayProp $lStatus $prop_name_cstr $location $rotation $font_to_apply $color] } }
                        if {$target_prop != "NULL"} {
                            if {$g_apply_flags(font)} { $target_prop SetFont [get_from_kv_list $src_prop_info "font"] }
                            if {$g_apply_flags(color)} { $target_prop SetColor [get_from_kv_list $src_prop_info "color"] }
                            if {$g_apply_flags(rotation)} { $target_prop SetRotation [get_from_kv_list $src_prop_info "rotation"] }
                            if {$g_apply_flags(display_type)} { $target_prop SetDisplayType [get_from_kv_list $src_prop_info "display_type"] }
                            if {$g_apply_flags(layout)} { $target_prop SetLocation [get_from_kv_list $src_prop_info "location"] }
                        }
                    }
                }
            }
        }
        if {$g_apply_flags(rotation) && [info exists g_copied_format(rotation)]} { if {![catch {$target_obj SetRotation $g_copied_format(rotation)}]} {} }
        if {$g_apply_flags(mirror) && [info exists g_copied_format(mirror)]} { if {![catch {$target_obj SetMirror $g_copied_format(mirror)}]} {} }
    }
    set msg "操作完成。\n\n成功应用到 $applied_count 个对象。"
    tk_messageBox -icon info -title "完成" -message $msg
}


# GUI 创建 / 管理
proc create_gui {} {
    global g_apply_flags
    if {[winfo exists .format_painter]} {destroy .format_painter}
    set w [toplevel .format_painter]; wm title $w "格式刷"; wm resizable $w 0 0; wm transient $w
    wm attributes .format_painter -topmost 1; wm protocol .format_painter WM_DELETE_WINDOW {destroy .format_painter}
    frame .format_painter.main_frame -padx 10 -pady 10; pack .format_painter.main_frame -fill both -expand true
    frame .format_painter.main_frame.buttons -pady 5
    button .format_painter.main_frame.buttons.copy -text "复制格式" -command copy_format
    button .format_painter.main_frame.buttons.paste -text "粘贴格式" -command paste_format
    button .format_painter.main_frame.buttons.close -text "关闭" -command {destroy .format_painter}
    pack .format_painter.main_frame.buttons.copy .format_painter.main_frame.buttons.paste -side left -padx 5 -pady 5
    pack .format_painter.main_frame.buttons.close -side right -padx 5 -pady 5
    pack .format_painter.main_frame.buttons -fill x
    labelframe .format_painter.main_frame.options -text "粘贴选项" -padx 5 -pady 5
    checkbutton .format_painter.main_frame.options.cb_font -text "字体" -variable g_apply_flags(font) -state disabled
    checkbutton .format_painter.main_frame.options.cb_color -text "颜色" -variable g_apply_flags(color) -state disabled
    checkbutton .format_painter.main_frame.options.cb_display_type -text "可见性" -variable g_apply_flags(display_type) -state disabled
    checkbutton .format_painter.main_frame.options.cb_layout -text "属性布局" -variable g_apply_flags(layout) -state disabled
    checkbutton .format_painter.main_frame.options.cb_rotation -text "旋转" -variable g_apply_flags(rotation) -state disabled
    checkbutton .format_painter.main_frame.options.cb_mirror -text "镜像" -variable g_apply_flags(mirror) -state disabled
    pack .format_painter.main_frame.options.cb_font .format_painter.main_frame.options.cb_color .format_painter.main_frame.options.cb_display_type .format_painter.main_frame.options.cb_layout .format_painter.main_frame.options.cb_rotation .format_painter.main_frame.options.cb_mirror -anchor w
    pack .format_painter.main_frame.options -fill x -pady 5
    labelframe .format_painter.main_frame.feedback -text "已复制格式信息" -padx 5 -pady 5
    label .format_painter.main_frame.feedback_label -text "源对象: 无" -justify left -wraplength 220
    pack .format_painter.main_frame.feedback_label -anchor w -fill x
    pack .format_painter.main_frame.feedback -fill x
    update; grab .format_painter; tkwait window .format_painter
}

# ===================================================================
#  OrCAD UI 集成函数
# ===================================================================
proc RMB_Copy_Format {} {
    if {![winfo exists .format_painter] || ![winfo ismapped .format_painter]} { LaunchFormatPainterGUI; after 100 { copy_format } } else { copy_format }
    raise .format_painter; focus .format_painter
}

proc Enable_Always {} { return 1 }

proc LaunchFormatPainterGUI {} {
    if {[winfo exists .format_painter] && [winfo ismapped .format_painter]} { raise .format_painter; focus .format_painter } else { create_gui }
}

proc RunFormatPainter {} {
    after idle create_gui
}

# ===================================================================
#  脚本主入口 - 右键菜单集成
# ===================================================================
proc RegisterFormatPainter_RMB {} {
    if {[catch {RegisterAction "---" "" "" "" "Schematic"}]} {}
    if {[catch {RegisterAction "formatBrush" "Enable_Always" "" "RMB_Copy_Format" "Schematic"} result]} {
        puts "ERROR: Failed to add 'formatBrush' to RMB menu. Reason: $result"
    }
}

after idle {
    RegisterFormatPainter_RMB
}

proc fp {} {
    RunFormatPainter
}