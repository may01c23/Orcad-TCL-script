################################################################################
# FormatPainterUtility.tcl
# 
# OrCAD Capture 格式刷工具
################################################################################

package require Tk 8.4
package provide FormatPainterUtility 1.0

# 隐藏主窗口
wm withdraw .

#===============================================================================
# 1. 命名空间与常量定义
#===============================================================================

namespace eval ::FormatPainterUtility {
    # 导出主函数
    namespace export show hide

    # 状态变量
    variable State
    array set State {
        WindowCreated 0
        WindowVisible 0
        CopiedFormatType -1
        CopiedDisplayProps {}
        CopiedRotation ""
        CopiedMirror ""
        IsTextMode 0
    }

    # 选项开关
    variable Options
    array set Options {
        Font 1
        Color 1
        Rotation 1
        Mirror 1
        DisplayType 1
        Layout 1
    }

    # UI 组件句柄
    variable Widgets
    array set Widgets {}

    # OrCAD 对象类型常量 (参考 DboBaseObject::ObjectTypeT)
    variable OBJ_TYPE
    array set OBJ_TYPE {
        BLOCK           12
        PART            13
        PORT            23
        POWER           37
        OFF_PAGE        38
        PART_PROP       39
        NET_ALIAS       49
        # 根据实际情况补充
    }

    # 类型 ID 到名称的映射
    variable TYPE_NAMES
    array set TYPE_NAMES {
        12 "Hierarchical Block"
        13 "Part Instance"
        23 "Hierarchical Port"
        37 "Power Symbol"
        38 "Off-Page Connector"
        39 "Part Property"
        49 "Net Alias"
		29 "Graphic Text"
    }
}

#===============================================================================
# 2. 核心辅助函数 (API Wrapper & Helpers)
#===============================================================================

# 安全执行 Dbo 操作，自动管理 DboState
# usage: WithDboState statusVar { body }
proc ::FormatPainterUtility::WithDboState {statusVarName body} {
    upvar 1 $statusVarName lStatus
    set lStatus [DboState]

    # 执行代码块，捕获错误
    set code [catch {uplevel 1 $body} result]

    # 务必删除 DboState 以防内存泄漏 (API 规则)
    catch {$lStatus -delete}

    if {$code == 1} {
        return -code error $result
    }
    return $result
}

# CString 转换辅助：Tcl String -> CString
# 返回的 CString 对象需要稍后手动删除，或者由 Capture 内部管理
# 这里我们简单封装，注意：如果在循环中大量使用，建议显式 Delete
proc ::FormatPainterUtility::MakeCString {tclStr} {
    return [DboTclHelper_sMakeCString $tclStr]
}

# CString 转换辅助：CString -> Tcl String
proc ::FormatPainterUtility::GetCString {cStrObj} {
    return [DboTclHelper_sGetConstCharPtr $cStrObj]
}

# 获取对象的显示名称 (RefDes 或 Name)
proc ::FormatPainterUtility::GetDisplayName {pObj} {
    set displayName "Unknown Object"
    ::FormatPainterUtility::WithDboState lStatus {
        set nameCStr [DboTclHelper_sMakeCString]

        # 尝试获取位号 (RefDes)
        if {![catch {$pObj GetReferenceDesignator $nameCStr} result] && $result != ""} {
             set displayName [DboTclHelper_sGetConstCharPtr $nameCStr]
        } else {
            # 尝试获取名称
            if {![catch {$pObj GetName $nameCStr}]} {
                set displayName [DboTclHelper_sGetConstCharPtr $nameCStr]
            }
        }
        # 释放内存
        DboTclHelper_sDeleteCString $nameCStr
    }
    return $displayName
}

# 检查对象是否具有特定方法
proc ::FormatPainterUtility::ObjectHasMethod {pObj method} {
    set hasMethod 1
    # 这里使用 try-catch 探测法
    ::FormatPainterUtility::WithDboState lStatus {
        if {[catch {$pObj $method $lStatus} _]} {
            set hasMethod 0
        }
    }
    return $hasMethod
}

# 获取对象的所有显示属性详细信息
proc ::FormatPainterUtility::CollectDisplayProps {pObj} {
    set propsList [list]

    ::FormatPainterUtility::WithDboState lStatus {
        # 检查是否支持迭代器
        if {[catch {$pObj NewDisplayPropsIter $lStatus} iter]} {
            return $propsList
        }

        # 预先创建 CString 缓冲区，避免在循环中反复创建销毁
        set nameCStr [DboTclHelper_sMakeCString]
        set lFont [DboTclHelper_sMakeLOGFONT]

        set pProp [$iter NextProp $lStatus]
        while {$pProp != "NULL"} {
            set propInfo [list]

            # 获取属性名
            $pProp GetName $nameCStr
            lappend propInfo "name" [DboTclHelper_sGetConstCharPtr $nameCStr]

            # 获取颜色
            lappend propInfo "color" [$pProp GetColor $lStatus]

            # 获取字体 (LOGFONT) - 注意：GetFont 会填充传入的 lFont 对象
            # 我们需要创建一个新的 LOGFONT 对象来存储结果，否则所有属性都会指向同一个 lFont 引用(如果直接存对象)
            # 或者我们需要把 lFont 的内容复制出来。
            # 但在 Tcl 中，DboTclHelper_sMakeLOGFONT 返回的是一个指针/句柄。
            # 如果我们想保存它，我们需要为每个属性创建一个新的，或者序列化它。
            # 简单起见，这里我们每次创建一个新的，但要注意这可能会有泄漏，除非有 DeleteLOGFONT。
            # 既然没有明确的 DeleteLOGFONT 文档，且 LOGFONT 结构较小，我们暂时保持原样，
            # 但为了安全，我们在循环内创建，循环外无法释放（除非存下来）。
            # 实际上，OrCAD Tcl API 中 LOGFONT 通常由系统管理或忽略泄漏。
            # 但为了严谨，我们尝试寻找 Delete 方法。如果没有，就只能这样。
            # 修正：为了避免 CString 泄漏，nameCStr 必须复用。
            
            # 关于 Font: 每次都 Make 一个新的给列表存着是必要的，因为后续 Paste 要用。
            # Paste 完后这些 Font 对象会泄漏，除非我们在 Paste 完后清理 State(CopiedDisplayProps)。
            # 这是一个更复杂的问题。暂时只修复 CString。
            
            set propFont [DboTclHelper_sMakeLOGFONT]
            $pProp GetFont $::DboLib_DEFAULT_FONT_PROPERTY $propFont
            lappend propInfo "font" $propFont

            # 获取旋转
            lappend propInfo "rotation" [$pProp GetRotation $lStatus]

            # 获取可见性 (DisplayType)
            lappend propInfo "display_type" [$pProp GetDisplayType $lStatus]

            # 获取位置 (相对位置)
            lappend propInfo "location" [$pProp GetLocation $lStatus]

            lappend propsList $propInfo

            # 下一个
            set pProp [$iter NextProp $lStatus]
        }
        
        # 释放缓冲区
        DboTclHelper_sDeleteCString $nameCStr
        # lFont 在这里没用上（上面循环里每次都新造了），所以删掉这个临时的
        # 假设没有 DeleteLOGFONT，就不删了，或者如果 API 支持 DeleteObject...
        # 通常 CString 是最主要的泄漏源。

        # 删除迭代器 (API 规则)
        delete_DboDisplayPropsIter $iter
    }

    return $propsList
}

#===============================================================================
# 3. 业务逻辑 (Copy/Paste Operations)
#===============================================================================

# 复制格式逻辑
proc ::FormatPainterUtility::PerformCopy {} {
    variable State
    variable Widgets
    variable Options
    variable TYPE_NAMES
    variable OBJ_TYPE

    # 检查窗口是否存在
    if {![winfo exists $Widgets(mainWindow)]} { return }

    set selectedObjs [GetSelectedObjects]
    if {[llength $selectedObjs] != 1} {
        tk_messageBox -parent $Widgets(mainWindow) -icon warning -title "选择错误" -message "请选中且仅选中一个对象作为源。"
        return
    }

    set pObj [lindex $selectedObjs 0]
    set objType [$pObj GetObjectType]

    # 检查类型支持
    if {![info exists TYPE_NAMES($objType)]} {
        tk_messageBox -parent $Widgets(mainWindow) -icon error -title "不支持" -message "不支持的对象类型 (ID: $objType)。"
        return
    }

    # 重置状态
    set State(CopiedFormatType) $objType
    set State(CopiedDisplayProps) [list]
    set State(CopiedRotation) ""
    set State(CopiedMirror) ""
    set feedbackList [list]

    set objName [GetDisplayName $pObj]

    ::FormatPainterUtility::WithDboState lStatus {
        # ---------------------------------------------------------
        # 情况 A: 纯文本对象或单个属性对象 (Net Alias, Part Property)
        # ---------------------------------------------------------
        if {$objType == $OBJ_TYPE(NET_ALIAS) || $objType == $OBJ_TYPE(PART_PROP)} {
            set State(IsTextMode) 1
            set propInfo [list]

            if {$objType == $OBJ_TYPE(NET_ALIAS)} {
                # Net Alias 只有 Font 和 Color
                lappend propInfo "font" [$pObj GetFont $lStatus]
                lappend propInfo "color" [$pObj GetColor $lStatus]
                lappend feedbackList "字体" "颜色"

                # UI 互锁：禁用不相关的选项
                set Options(DisplayType) 0
                set Options(Layout) 0
                set Options(Font) 1
                set Options(Color) 1
            } else {
                # Property 对象
                set lFont [DboTclHelper_sMakeLOGFONT]
                $pObj GetFont $::DboLib_DEFAULT_FONT_PROPERTY $lFont
                lappend propInfo "font" $lFont
                lappend propInfo "color" [$pObj GetColor $lStatus]

                if {$objType == $OBJ_TYPE(PART_PROP)} {
                    lappend propInfo "display_type" [$pObj GetDisplayType $lStatus]
                    lappend feedbackList "可见性"
                    set Options(DisplayType) 1
                } else {
                    set Options(DisplayType) 0
                }

                lappend feedbackList "字体" "颜色"
                set Options(Layout) 0
                set Options(Font) 1
                set Options(Color) 1
            }
            lappend State(CopiedDisplayProps) $propInfo

        } else {
            # ---------------------------------------------------------
            # 情况 B: 容器对象 (Part, Block, etc.)
            # ---------------------------------------------------------
            set State(IsTextMode) 0
            if {[ObjectHasMethod $pObj "NewDisplayPropsIter"]} {
                set State(CopiedDisplayProps) [CollectDisplayProps $pObj]
                if {[llength $State(CopiedDisplayProps)] > 0} {
                    lappend feedbackList "字体" "颜色" "可见性" "位置"
                    # 启用所有相关选项
                    set Options(Font) 1
                    set Options(Color) 1
                    set Options(DisplayType) 1
                    set Options(Layout) 1
                }
            }
        }

        # ---------------------------------------------------------
        # 通用属性：旋转与镜像
        # ---------------------------------------------------------
        if {[ObjectHasMethod $pObj "GetRotation"]} {
            set State(CopiedRotation) [$pObj GetRotation $lStatus]
            lappend feedbackList "旋转"
            set Options(Rotation) 1
        } else {
            set Options(Rotation) 0
        }

        if {[ObjectHasMethod $pObj "GetMirror"]} {
            set State(CopiedMirror) [$pObj GetMirror $lStatus]
            lappend feedbackList "镜像"
            set Options(Mirror) 1
        } else {
            set Options(Mirror) 0
        }
    }

    # puts $State(CopiedDisplayProps)

    # 更新 UI 反馈
    ::FormatPainterUtility::UpdateUIState
    set feedbackStr [join [lsort -unique $feedbackList] ", "]
    $Widgets(feedbackLabel) configure -text "已复制源: $objName";# \n包含: $feedbackStr"
}

# 粘贴格式逻辑
proc ::FormatPainterUtility::PerformPaste {} {
    variable State
    variable Widgets
    variable Options
    variable TYPE_NAMES
	variable CONST
	set DO_NOT_DISPLAY 0

    # 1. 检查是否已复制源格式
    if {$State(CopiedFormatType) == -1} {
        tk_messageBox -parent $Widgets(mainWindow) -icon warning -title "警告" -message "请先复制源对象格式。"
        return
    }

    # 2. 获取选中的对象
    set selectedObjs [GetSelectedObjects]
    set totalSelected [llength $selectedObjs]
    if {$totalSelected == 0} { return }

    set successCount 0
    set skippedCount 0

    # 获取源类型的名称，用于错误报告
    set srcTypeName "未知类型"
    if {[info exists TYPE_NAMES($State(CopiedFormatType))]} {
        set srcTypeName $TYPE_NAMES($State(CopiedFormatType))
    }
    # 2. 批量应用 (使用 catch 确保单个失败不中断整体)
    # 使用 WithDboState 自动管理状态对象
    ::FormatPainterUtility::WithDboState lStatus {
        foreach target_obj $selectedObjs {
            # --- 单个对象类型检查 ---
            set targetType [$target_obj GetObjectType]

            if {$targetType != $State(CopiedFormatType)} {
                # 类型不匹配，计数并跳过
                incr skippedCount
                continue
            }

            # --- 类型匹配，执行应用逻辑 ---
            if {[catch {
                # 准备源属性数据数组
                # 取出第一个属性块（对于 TextMode 只有一个，对于 Part 模式这是第一个）
                set srcPropList [lindex $State(CopiedDisplayProps) 0]

                # 清空并设置临时数组，用来模拟 dict
                unset -nocomplain srcPropArr
                array set srcPropArr $srcPropList
                # A. 处理单文本/属性对象
                if {$State(IsTextMode)} {
					set srcPropList [lindex $State(CopiedDisplayProps) 0]
                    if {$Options(Font) && [info exists srcPropArr(font)]} {

                        $target_obj SetFont $srcPropArr(font)

                    }
                    if {$Options(Color) && [info exists srcPropArr(color)]} {
                        $target_obj SetColor $srcPropArr(color)
                    }
                    if {$Options(DisplayType) && [info exists srcPropArr(display_type)]} {
                        # 只有部分对象支持 SetDisplayType
                        if {[::FormatPainterUtility::ObjectHasMethod $target_obj "SetDisplayType"]} {
                            $target_obj SetDisplayType $srcPropArr(display_type)
                        }
                    }
                # B. 容器对象模式 (Part, Block 等) 
                } elseif {[::FormatPainterUtility::ObjectHasMethod $target_obj "NewDisplayPropsIter"]} {
                    # 1. 预处理：解析源对象中那些是“可见”的属性名
                    set source_visible_prop_names [list]
                    foreach srcPropList $State(CopiedDisplayProps) {
                        unset -nocomplain srcPropArr
                        array set srcPropArr $srcPropList

                        # 如果源属性是显示的，加入白名单
                        if {[info exists srcPropArr(display_type)] && $srcPropArr(display_type) != $DO_NOT_DISPLAY} {
                            lappend source_visible_prop_names $srcPropArr(name)
                        }
                    }

                    # 2. 清洗步骤：如果勾选了“可见性”，则隐藏目标对象上那些“不在源可见列表”中的属性
                    if {$Options(DisplayType)} {
                        set target_props_iter [$target_obj NewDisplayPropsIter $lStatus]
                        set prop_to_check [$target_props_iter NextProp $lStatus]
                        
                        # 复用 CString
                        set prop_name_cstr [MakeCString ""]

                        while {$prop_to_check != "NULL"} {
                            $prop_to_check GetName $prop_name_cstr
                            set prop_name [GetCString $prop_name_cstr]

                            # 获取当前显示状态
                            set current_disp_type [$prop_to_check GetDisplayType $lStatus]

                            # 如果当前是显示的，但不在源的可见列表中，则强制隐藏
                            if {$current_disp_type != $DO_NOT_DISPLAY && \
                                [lsearch -exact $source_visible_prop_names $prop_name] == -1} {
                                $prop_to_check SetDisplayType $DO_NOT_DISPLAY 
                            }
                            set prop_to_check [$target_props_iter NextProp $lStatus]
                        }
                        DboTclHelper_sDeleteCString $prop_name_cstr
                        delete_DboDisplayPropsIter $target_props_iter
                    }

                    # 3. 构建目标对象的有效属性白名单 (Target Matchable Props)
                    #    逻辑：只有目标对象本身拥有的属性 (Effective Props)，并且该属性也在源对象的可见列表中
                    #    (或者如果不刷可见性，则只要源对象有这个属性配置即可)
                    set target_visible_prop_names [list]
                    set target_eff_props_iter [$target_obj NewEffectivePropsIter $lStatus]

                    # 准备迭代器所需的临时变量
                    set prop_name_cstr [MakeCString ""]
                    set prop_value_cstr [MakeCString ""]
                    set prop_type_cstr [DboTclHelper_sMakeDboValueType]
                    set prop_edit_cstr [DboTclHelper_sMakeInt]

                    set iterStatus [$target_eff_props_iter NextEffectiveProp $prop_name_cstr $prop_value_cstr $prop_type_cstr $prop_edit_cstr]

                    while {[$iterStatus OK] == 1} {
                        set prop_name [GetCString $prop_name_cstr]

                        # 如果勾选了可见性，我们只关心源中可见的属性
                        if {$Options(DisplayType)} {
                            if {[lsearch -exact $source_visible_prop_names $prop_name] != -1} {
                                lappend target_visible_prop_names $prop_name
                            }
                        } else {
                            # 如果不勾选可见性，我们允许匹配所有属性
                            lappend target_visible_prop_names $prop_name
                        }

                        set iterStatus [$target_eff_props_iter NextEffectiveProp $prop_name_cstr $prop_value_cstr $prop_type_cstr $prop_edit_cstr]
                    }
                    
                    # 清理 CString
                    DboTclHelper_sDeleteCString $prop_name_cstr
                    DboTclHelper_sDeleteCString $prop_value_cstr
                    # 注意：DboValueType 和 Int 指针通常不需要 Delete，或者是 DeleteInt/DeleteDboValueType (如果存在)
                    
                    delete_DboEffectivePropsIter $target_eff_props_iter

                    # 4. 核心循环：遍历源格式，应用到目标
                    if {[llength $State(CopiedDisplayProps)] > 0} {
                        foreach srcPropList $State(CopiedDisplayProps) {
                            unset -nocomplain srcPropArr
                            array set srcPropArr $srcPropList

                            set prop_name $srcPropArr(name)

                            # 只有在步骤3中确认目标对象拥有该属性，才进行操作
                            if {[lsearch -exact $target_visible_prop_names $prop_name] != -1} {
                                # 修正：不能复用 CString 并调用 Set (可能不存在或有问题)，
                                # 必须每次创建新的 CString 对象
                                set prop_name_cstr [MakeCString $prop_name]
                                set target_prop [$target_obj GetDisplayProp $prop_name_cstr $lStatus]

                                if {$target_prop != "NULL"} {
                                    # --- 情况 4.1: 属性已存在显示实例 (Update) ---

                                    if {$Options(Font) && [info exists srcPropArr(font)]} {
                                        $target_prop SetFont $srcPropArr(font)
                                    }
                                    if {$Options(Color) && [info exists srcPropArr(color)]} {
                                        $target_prop SetColor $srcPropArr(color)
                                    }
                                    if {$Options(Rotation) && [info exists srcPropArr(rotation)]} {
                                        $target_prop SetRotation $srcPropArr(rotation)
                                    }
                                    if {$Options(DisplayType) && [info exists srcPropArr(display_type)]} {
                                        $target_prop SetDisplayType $srcPropArr(display_type)
                                    }
                                    # 这里是你要求的 location 更新
                                    if {$Options(Layout) && [info exists srcPropArr(location)]} {
                                        $target_prop SetLocation $srcPropArr(location)
                                    }

                                } else {
                                    # --- 情况 4.2: 属性存在但未显示 (Create) ---
                                    # 只有当源属性要求显示，且我们开启了可见性选项时，才创建新的 DisplayProp

                                    set src_disp_type $srcPropArr(display_type)

                                    if {$Options(DisplayType) && [info exists srcPropArr(display_type)] && $src_disp_type != $DO_NOT_DISPLAY} {
                                        set prop_location $srcPropArr(location)
                                        set prop_rotation $srcPropArr(rotation)
                                        set prop_font     $srcPropArr(font)
                                        set prop_color    $srcPropArr(color)

                                        # NewDisplayProp 需要传 lStatus
                                        set pNewDispProp [$target_obj NewDisplayProp $lStatus $prop_name_cstr $prop_location $prop_rotation $prop_font $prop_color]

                                        if {$pNewDispProp != "NULL"} {
                                            $pNewDispProp SetDisplayType $src_disp_type
                                        }
                                    }
                                }
                                # 每次循环结束释放 CString
                                DboTclHelper_sDeleteCString $prop_name_cstr
                            }
                        }
                    }
                }

                # ===========================================================
                # C. 通用变换 (Part/Alias 本身的旋转镜像)
                # ===========================================================
                if {$Options(Rotation) && $State(CopiedRotation) != ""} {
                    $target_obj SetRotation $State(CopiedRotation)
                }
                if {$Options(Mirror) && $State(CopiedMirror) != ""} {
                    $target_obj SetMirror $State(CopiedMirror)
                }

                incr successCount

            } err]} {
				# 捕获应用过程中的意外错误，不打断循环
                set errCStr [MakeCString "FormatPainter Paste Error on object $target_obj: $err"]
                DboState_WriteToSessionLog $errCStr
                DboTclHelper_sDeleteCString $errCStr
            }
        }
    } ;# End WithDboState

    # 刷新画布
    catch {ZoomRedraw}
    set msg "操作完成。\n\n"
    append msg "总共选择: $totalSelected\n"
    append msg "成功应用: $successCount\n"

    if {$skippedCount > 0} {
        append msg "--------------------\n"
        append msg "跳过对象: $skippedCount (类型与源不匹配)"
        tk_messageBox -parent $Widgets(mainWindow) -icon warning -title "部分完成" -message $msg

    } else {
        # 可选：状态栏提示而不是弹窗，体验更好
		# DboState_WriteToSessionLog [MakeCString "Format Painter applied to $successCount objects."]
		tk_messageBox -parent $Widgets(mainWindow) -icon info -title "完成" -message "成功应用到 $successCount 个对象。"
    }
}

# 更新复选框状态 (Enable/Disable)
proc ::FormatPainterUtility::UpdateUIState {} {
    variable Widgets
    variable Options

    foreach opt {Font Color Rotation Mirror DisplayType Layout} {
        set key [string tolower $opt]
        # 实际上在 Copy 时已经设置了 Options 的值 (0 或 1)
        # 这里我们根据 Options 的值来 Enable/Disable 控件交互？
        # 或者仅仅是根据 Copy 的结果决定哪些 Checkbox 是可用的 (State = normal/disabled)

        # 逻辑修正：如果源对象有这个属性 (Options=1)，则 Checkbox 状态为 Normal 且默认选中
        # 如果源对象没有 (Options=0)，则 Disabled

        if {[info exists Widgets(${key}Check)]} {
            if {$Options($opt)} {
                $Widgets(${key}Check) configure -state normal
                # 默认勾选
                # set ::FormatPainterUtility::Options($opt) 1 
            } else {
                $Widgets(${key}Check) configure -state disabled
                # 取消勾选
                # set ::FormatPainterUtility::Options($opt) 0
            }
        }
    }
}

#===============================================================================
# 4. UI 构建模块 (修改版：大字体 + 按钮并排)
#===============================================================================

proc ::FormatPainterUtility::CreateUI {} {
    variable Widgets
    variable State

    if {[winfo exists .formatPainterDlg]} {
        wm deiconify .formatPainterDlg
        raise .formatPainterDlg
        return
    }

    # --- 1. 定义样式配置 (修改这里可以改变大小) ---
    # Windows默认通常是 9号，改大到 11 或 12 会显著变大
    set uiFont "微软雅黑 9" 
    set btnFont "微软雅黑 9 bold"
    # 增加内边距让界面更宽松
    set padX 10
    set padY 2
    
    set dlg [toplevel .formatPainterDlg]
    wm title $dlg "OrCAD Format Painter"
    # 允许调整窗口大小 (可选，如果界面太大可能需要拉伸)
    wm resizable $dlg 1 1 
    wm attributes $dlg -topmost 1
    
    wm protocol $dlg WM_DELETE_WINDOW {::FormatPainterUtility::hide}
    
    set Widgets(mainWindow) $dlg

    # --- 主容器 ---
    # 增加 padx/pady 让整体边缘留白更多
    set mainFr [frame $dlg.fr -padx 20 -pady 20]
    pack $mainFr -fill both -expand 1

    # --- A. 按钮区 (修改为并排) ---
    set btnFr [frame $mainFr.btn -pady 5]
    pack $btnFr -fill x

    # 创建按钮，应用大字体
    button $btnFr.copy -text "1. 复制格式" -command ::FormatPainterUtility::PerformCopy -font $btnFont -height 2
    button $btnFr.paste -text "2. 粘贴格式" -command ::FormatPainterUtility::PerformPaste -font $btnFont -height 2
    
    # [修改点] 将 side 改为 left，并添加 expand/fill 让它们并排且等宽
    pack $btnFr.copy -side left -expand 1 -fill x -padx 5
    pack $btnFr.paste -side left -expand 1 -fill x -padx 5

    # --- B. 选项区 ---
    set optFr [labelframe $mainFr.opt -text "应用选项" -padx $padX -pady $padY -font $btnFont]
    pack $optFr -fill x -pady 10

    # 使用 grid 布局
    set r 0; set c 0
    foreach {opt label} {
        Font "字体" 
        Color "颜色" 
        DisplayType "可见性" 
        Layout "位置" 
        Rotation "旋转" 
        Mirror "镜像"
    } {
        set key [string tolower $opt]
        # 应用大字体
        checkbutton $optFr.cb_$key -text " $label" -variable ::FormatPainterUtility::Options($opt) \
            -state disabled -font $uiFont
        
        # sticky w (靠左对齐), padx 增加横向间距
        grid $optFr.cb_$key -row $r -column $c -sticky w -padx 10 -pady 5
        set Widgets(${key}Check) $optFr.cb_$key
        
        # 两列布局逻辑
        incr c
        if {$c > 1} { set c 0; incr r }
    }
    
    # 让Grid列自适应
    grid columnconfigure $optFr 0 -weight 1
    grid columnconfigure $optFr 1 -weight 1

    # --- C. 反馈区 ---
    set fbFr [labelframe $mainFr.fb -text "当前格式源" -fg blue -padx $padX -pady $padY -font $btnFont]
    pack $fbFr -fill x -pady 10
    
    # -wraplength 调大，防止大字体下换行太频繁
    label $fbFr.lbl -text "未选择源对象 (请先复制)" -wraplength 350 -justify left -anchor w -font $uiFont
    pack $fbFr.lbl -fill x -anchor w
    set Widgets(feedbackLabel) $fbFr.lbl

    # 绑定快捷键
    # bind $dlg <Control-c> ::FormatPainterUtility::PerformCopy
    # bind $dlg <Control-v> ::FormatPainterUtility::PerformPaste
}

#===============================================================================
# 5. 公共入口
#===============================================================================

proc ::FormatPainterUtility::show {} {
    variable Widgets
    CreateUI
    # 尝试设置为模态窗口，但要注意 catch 保护
    if {[catch {grab $Widgets(mainWindow)} err]} {
        DboState_WriteToSessionLog [MakeCString "Warning: Could not grab window: $err"]
    }
}

proc ::FormatPainterUtility::hide {} {
    variable Widgets
    if {[info exists Widgets(mainWindow)] && [winfo exists $Widgets(mainWindow)]} {
        grab release $Widgets(mainWindow)
        destroy $Widgets(mainWindow)
    }
}

# puts "Run '::FormatPainterUtility::show' to start."
