################################################################################
# NCMarkerUtility.tcl
# 
# OrCAD Capture 一键 DNI/NC 标记工具
################################################################################

package require Tk 8.4
package provide NCMarkerUtility 1.0

# 隐藏主窗口
wm withdraw .

#===============================================================================
# 1. 全局配置与常量定义
#===============================================================================

namespace eval ::NCMarkerUtility {
    variable widgets
    variable options
    variable state
    variable constants
    
    array set widgets {}
    array set options {
        DNI 1
        modifyValue 1
        bounding 1
        propName "DNI"
        propValue "NC"
        propSuffix "/NC"
    }
    array set state {
        window_created 0
        window_visible 0
    }
    
    # 定义常量
    array set constants {
        OBJ_TYPE_BLOCK 12
        OBJ_TYPE_PART  13
        OBJ_TYPE_GRAPHIC_BOX 4
        
        LINE_STYLE_DASH     3
        LINE_WIDTH_THIN     0
        FILL_STYLE_Cross    2
        HATCH_STYLE_DIAG    3
        COLOR_RED           3 
        
        DISP_VALUE_TYPE     5
        DISP_NAME_AND_VALUE 0
        lx 1
        ly 1
    }
}

#===============================================================================
# 2. 核心逻辑与工具函数
#===============================================================================

namespace eval ::NCMarkerUtility::Utils {
    namespace export drawBoundingBox getObjProps applyToSelection removeSelection
}

# 检查对象能力
proc ::NCMarkerUtility::Utils::objectHasCapability {obj capability} {
    if {[llength [info commands $obj]] > 0} {
        return [expr {![catch {$obj $capability} err]}]
    }
    return 0
}

# 绘制围框
proc ::NCMarkerUtility::Utils::drawBoundingBox {obj} {
    variable ::NCMarkerUtility::constants
    
    set lStatus [DboState]
    set lBBox [$obj GetOffsetBoundingBox $lStatus]
    
    if {[$lStatus OK]} {
        set lPage [$obj GetOwner]
        
        set lLocation [$obj GetLocation $lStatus]
        set lRotation [$obj GetRotation $lStatus]
        
        # 创建 Box
        set lGraphicObj [$lPage NewGraphicBoxInst $lStatus $lBBox $lLocation $lRotation]
        
        if {[$lStatus OK] && $lGraphicObj != "NULL"} {
            $lGraphicObj SetLineStyle $constants(LINE_STYLE_DASH)
            $lGraphicObj SetLineWidth $constants(LINE_WIDTH_THIN)
            $lGraphicObj SetFillStyle $constants(FILL_STYLE_Cross)
            $lGraphicObj SetHatchStyle $constants(HATCH_STYLE_DIAG)
            $lGraphicObj SetColor $constants(COLOR_RED)
        }
        
        # 清理
        if {$lLocation != "NULL"} { DboTclHelper_sDeleteCPoint $lLocation }
        # Rotation 通常是 int
    }
    
    catch {DboTclHelper_sDeleteCRect $lBBox}
    $lStatus -delete
}

# 获取对象属性
proc ::NCMarkerUtility::Utils::getObjProps {obj} {
    # 初始化默认列表
    set prop_list [list "font" "" "color" 0 "rotation" 0]
    
    set lStatus [DboState]
    
    # 检查是否支持 NewDisplayPropsIter
    if {[objectHasCapability $obj "NewDisplayPropsIter"]} {
        set props_iter [$obj NewDisplayPropsIter $lStatus]
        
        if {[$lStatus OK]} {
            set prop [$props_iter NextProp $lStatus]
            while {$prop != "NULL"} {
                set lNameStr [DboTclHelper_sMakeCString]
                $prop GetName $lNameStr
                set propName [DboTclHelper_sGetConstCharPtr $lNameStr]
                DboTclHelper_sDeleteCString $lNameStr
                
                # 寻找参考属性 "Part Reference"
                if {$propName == "Part Reference"} {
                    # 更新列表中的值
                    array set tempArr $prop_list
                    
                    set tempArr(color) [$prop GetColor $lStatus]
                    set tempArr(rotation) [$prop GetRotation $lStatus]
                    
                    set lFont [DboTclHelper_sMakeLOGFONT]
                    $prop GetFont $::DboLib_DEFAULT_FONT_PROPERTY $lFont
                    set tempArr(font) $lFont
                    
                    set prop_list [array get tempArr]
                    break
                }
                set prop [$props_iter NextProp $lStatus]
            }
        }
        delete_DboDisplayPropsIter $props_iter
    }
    
    $lStatus -delete
    return $prop_list
}

# 应用主逻辑
proc ::NCMarkerUtility::Utils::applyToSelection {} {
    variable ::NCMarkerUtility::options
    variable ::NCMarkerUtility::constants
    
    set selected_objects [GetSelectedObjects]
    if {[llength $selected_objects] == 0} {
        tk_messageBox -icon error -title "错误" -message "请先选择元件(Part)。"
        return
    }
    
    set lStatus [DboState]
    set count 0
    
    foreach obj $selected_objects {
        set type_id [$obj GetObjectType]
        # 仅处理 Part(13) 或 Block(12)
        if {$type_id != $constants(OBJ_TYPE_PART) && $type_id != $constants(OBJ_TYPE_BLOCK)} {
            continue
        }
        
        incr count
        
        # 获取样式信息 (返回的是 Key-Value List)
        set styleList [getObjProps $obj]
        array set styleArr $styleList
        
        # 绘制围框
        if {$options(bounding)} {
            drawBoundingBox $obj
        }
        
        # 1. 处理 Value 属性后缀 (v1.3 新增, v1.4 增加开关)
        if {$options(modifyValue) && $options(propSuffix) != ""} {
            set valPropName [DboTclHelper_sMakeCString "Value"]
            set lVal [DboTclHelper_sMakeCString]
            $obj GetEffectivePropStringValue $valPropName $lVal
            set currentVal [DboTclHelper_sGetConstCharPtr $lVal]
            
            # 检查是否已包含后缀
            # 修复: 当 currentVal 比 propSuffix 短时，string last 返回 -1，
            # 而 length 差值也是负数，导致错误判断为已包含。
            # 需要确保 expectedIndex >= 0
            set suffixLen [string length $options(propSuffix)]
            set valLen [string length $currentVal]
            set expectedIndex [expr {$valLen - $suffixLen}]
            
            if {$expectedIndex < 0 || [string last $options(propSuffix) $currentVal] != $expectedIndex} {
                set newVal "${currentVal}$options(propSuffix)"
                set newValStr [DboTclHelper_sMakeCString $newVal]
                $obj SetEffectivePropStringValue $valPropName $newValStr
                DboTclHelper_sDeleteCString $newValStr
            }
            
            DboTclHelper_sDeleteCString $valPropName
            DboTclHelper_sDeleteCString $lVal
        }
        
        # 2. 添加/更新 DNI 属性
        if {$options(DNI) && $options(propName) != ""} {
            set foundProp "NULL"
            set iter [$obj NewDisplayPropsIter $lStatus]
            set prop [$iter NextProp $lStatus]
            
            # 查找现有属性
            while {$prop != "NULL"} {
                set nameStr [DboTclHelper_sMakeCString]
                $prop GetName $nameStr
                if {[DboTclHelper_sGetConstCharPtr $nameStr] == $options(propName)} {
                    set foundProp $prop
                    DboTclHelper_sDeleteCString $nameStr
                    break
                }
                DboTclHelper_sDeleteCString $nameStr
                set prop [$iter NextProp $lStatus]
            }
            delete_DboDisplayPropsIter $iter
            
            set propName [DboTclHelper_sMakeCString $options(propName)]
            set propVal [DboTclHelper_sMakeCString $options(propValue)]
            set dispLoc [DboTclHelper_sMakeCPoint $constants(lx) $constants(ly)]

            # 如果存在，更新值
            if {$propName != "NULL" && $foundProp != "NULL"} {
                $obj SetEffectivePropStringValue $propName $propVal
                $foundProp SetDisplayType $constants(DISP_VALUE_TYPE)
                $foundProp SetLocation $dispLoc
            } else {
                # 如果不存在，创建新属性
                $obj SetEffectivePropStringValue $propName $propVal
                
                # 获取样式
                set font $styleArr(font)
                set fontCreated 0
                if {$font == ""} { 
                    set font [DboTclHelper_sMakeLOGFONT] 
                    set fontCreated 1
                }
                set color $styleArr(color)
                set rot $styleArr(rotation)
                
                set newProp [$obj NewDisplayProp $lStatus $propName $dispLoc $rot $font $color]
                
                if {[$lStatus OK] && $newProp != "NULL"} {
                    $newProp SetDisplayType $constants(DISP_VALUE_TYPE)
                }
                
                
                if {$fontCreated} { catch {DboTclHelper_sDeleteLOGFONT $font} }
            }
            
            # 清理
            DboTclHelper_sDeleteCString $propName
            DboTclHelper_sDeleteCString $propVal
            DboTclHelper_sDeleteCPoint $dispLoc
        }
        
        # 清理 styleArr 中的 font
        if {[info exists styleArr(font)] && $styleArr(font) != ""} {
             catch {DboTclHelper_sDeleteLOGFONT $styleArr(font)}
        }
        unset styleArr
    }
    
    $lStatus -delete
    
    # 刷新屏幕
    catch {ZoomRedraw}
    
    tk_messageBox -icon info -title "完成" -message "已处理 $count 个对象。"
}

# 移除标记逻辑
proc ::NCMarkerUtility::Utils::removeSelection {} {
    variable ::NCMarkerUtility::options
    variable ::NCMarkerUtility::constants
    
    set selected_objects [GetSelectedObjects]
    if {[llength $selected_objects] == 0} {
        tk_messageBox -icon error -title "错误" -message "请先选择元件(Part)。"
        return
    }
    
    set lStatus [DboState]
    set count 0
    
    foreach obj $selected_objects {
        set type_id [$obj GetObjectType]
        if {$type_id != $constants(OBJ_TYPE_PART) && $type_id != $constants(OBJ_TYPE_BLOCK)} {
            continue
        }
        incr count
        
        # 1. 移除 Value 后缀 (v1.4 增加开关)
        if {$options(modifyValue) && $options(propSuffix) != ""} {
            set valPropName [DboTclHelper_sMakeCString "Value"]
            set lVal [DboTclHelper_sMakeCString]
            $obj GetEffectivePropStringValue $valPropName $lVal
            set currentVal [DboTclHelper_sGetConstCharPtr $lVal]
            
            # 检查是否以后缀结尾
            if {[string last $options(propSuffix) $currentVal] == [expr {[string length $currentVal] - [string length $options(propSuffix)]}] && [string length $currentVal] > [string length $options(propSuffix)]} {
                set newVal [string range $currentVal 0 [expr {[string length $currentVal] - [string length $options(propSuffix)] - 1}]]
                set newValStr [DboTclHelper_sMakeCString $newVal]
                $obj SetEffectivePropStringValue $valPropName $newValStr
                DboTclHelper_sDeleteCString $newValStr
            }
            
            DboTclHelper_sDeleteCString $valPropName
            DboTclHelper_sDeleteCString $lVal
        }
        
        # 2. 移除 DNI 属性 (改为清空值并隐藏)
        set propName [DboTclHelper_sMakeCString $options(propName)]
        set emptyVal [DboTclHelper_sMakeCString ""]
        
        # 设置值为空
        $obj SetEffectivePropStringValue $propName $emptyVal
        
        # 尝试隐藏 DisplayProp (如果存在)
        set iter [$obj NewDisplayPropsIter $lStatus]
        set prop [$iter NextProp $lStatus]
        while {$prop != "NULL"} {
            set nameStr [DboTclHelper_sMakeCString]
            $prop GetName $nameStr
            if {[DboTclHelper_sGetConstCharPtr $nameStr] == $options(propName)} {
                $prop SetDisplayType 0
                DboTclHelper_sDeleteCString $nameStr
                break
            }
            
            DboTclHelper_sDeleteCString $nameStr
            set prop [$iter NextProp $lStatus]
        }
        delete_DboDisplayPropsIter $iter
        
        DboTclHelper_sDeleteCString $propName
        DboTclHelper_sDeleteCString $emptyVal
    }
    
    $lStatus -delete
    catch {ZoomRedraw}
    tk_messageBox -icon info -title "完成" -message "已移除 $count 个对象的标记。"
}

#===============================================================================
# 3. UI 界面 (Tk 8.4 Standard Widgets)
#===============================================================================

namespace eval ::NCMarkerUtility::UI {
    namespace export createWindow
}

proc ::NCMarkerUtility::UI::createWindow {} {
    variable ::NCMarkerUtility::widgets
    variable ::NCMarkerUtility::options

    set winName .nc_marker_util
    if {[winfo exists $winName]} {
        destroy $winName
    }
    
    toplevel $winName
    wm title $winName "NC标记工具"
    wm attributes $winName -topmost 1
    
    # 定义字体
    set uiFont "微软雅黑 9" 
    
    # 使用标准 frame
    set mainFrame [frame $winName.main]
    pack $mainFrame -fill both -expand true -padx 20 -pady 20
    
    # 1. 按钮区域 (置顶)
    set btnFrame [frame $mainFrame.btn]
    pack $btnFrame -fill x -pady 10
    
    button $btnFrame.apply -text "执行标记" -command ::NCMarkerUtility::Utils::applyToSelection -width 15 -font $uiFont
    button $btnFrame.remove -text "移除标记" -command ::NCMarkerUtility::Utils::removeSelection -width 15 -font $uiFont
    button $btnFrame.close -text "关闭" -command "destroy $winName" -width 10 -font $uiFont
    
    pack $btnFrame.apply -side left -padx 10
    pack $btnFrame.remove -side left -padx 10
    pack $btnFrame.close -side right -padx 10
    
    # 2. 选项区域 (Labelframe)
    set optFrame [labelframe $mainFrame.opt -text "配置" -padx 10 -pady 10 -font $uiFont]
    pack $optFrame -fill x -pady 10
    
    # 属性配置
    set propFrame [frame $optFrame.prop]
    pack $propFrame -fill x -pady 5
    
    label $propFrame.lbl_name -text "属性名:" -width 8 -font $uiFont
    entry $propFrame.ent_name -textvariable ::NCMarkerUtility::options(propName) -width 10 -font $uiFont
    
    label $propFrame.lbl_val -text "值:" -width 4 -font $uiFont
    entry $propFrame.ent_val -textvariable ::NCMarkerUtility::options(propValue) -width 10 -font $uiFont
    
    label $propFrame.lbl_suf -text "Value后缀:" -width 10 -font $uiFont
    entry $propFrame.ent_suf -textvariable ::NCMarkerUtility::options(propSuffix) -width 8 -font $uiFont
    
    pack $propFrame.lbl_name -side left
    pack $propFrame.ent_name -side left -padx 5
    pack $propFrame.lbl_val -side left
    pack $propFrame.ent_val -side left -padx 5
    pack $propFrame.lbl_suf -side left
    pack $propFrame.ent_suf -side left -padx 5
    
    # Checkbuttons
    checkbutton $optFrame.cb_dni -text "添加/更新属性 (DNI)" \
        -variable ::NCMarkerUtility::options(DNI) -font $uiFont
        
    checkbutton $optFrame.cb_val -text "修改 Value 后缀" \
        -variable ::NCMarkerUtility::options(modifyValue) -font $uiFont
        
    checkbutton $optFrame.cb_box -text "绘制虚线围框" \
        -variable ::NCMarkerUtility::options(bounding) -font $uiFont
        
    pack $optFrame.cb_dni -anchor w -pady 2
    pack $optFrame.cb_val -anchor w -pady 2
    pack $optFrame.cb_box -anchor w -pady 2
}

#===============================================================================
# 4. 启动入口
#===============================================================================

proc ::NCMarkerUtility::show {} {
    ::NCMarkerUtility::UI::createWindow
}

# puts "Run '::NCMarkerUtility::show' to start."