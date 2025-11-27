################################################################################
# NetPlacementUtility.tcl
# 
# OrCAD Capture 网络标号、Off-Page 批量放置与引脚网络生成工具
################################################################################

package require Tk 8.4
package provide NetPlacementUtility 1.0

# 隐藏主窗口，只显示配置对话框
wm withdraw .

namespace eval ::NetPlacementUtility {
    variable widgets
    variable options
    variable state
    
    array set widgets {}
    
    # --------------------------------------------------------------------------
    # 路径配置优化：使用 OrCAD 环境变量定位默认库
    # 参考: 文档 4.2 节中使用 $env(CDS_INST_DIR) 定位安装目录 
    # --------------------------------------------------------------------------
    if {[info exists env(CDS_INST_DIR)]} {
        set installDir [file normalize $env(CDS_INST_DIR)]
        set defaultLibPath [file join $installDir "tools" "capture" "library" "capsym.olb"]
    } else {
        # 回退方案：基于当前可执行文件推断
        set exePath [info nameofexecutable]
        set binDir [file dirname $exePath]
        set defaultLibPath [file join $binDir "library" "capsym.olb"]
    }
    set defaultLibPath [file normalize $defaultLibPath]

    # 默认配置参数
    array set options {
        operationType      "nets"
        generationMethod   "array"
        offpageDirection   "in"
        prefix             ""
        startNum           0
        endNum             7
        suffix             ""
    }
    set options(libPath) $defaultLibPath
             
    array set state {
        windowCreated 0
        windowVisible 0
    }
}

#===============================================================================
# Utils: 工具函数与逻辑实现
#===============================================================================
namespace eval ::NetPlacementUtility::Utils {
    namespace export placeNetsArray placeNetsList placeOffpageInArray placeOffpageInList \
                     placeOffpageOutArray placeOffpageOutList getSelectedPinNames placePinNetsArray \
                     getPlacementOrigin
}

# 获取鼠标最后点击位置作为放置原点
# API 参考: GetLastMouseClickPointOnPage 
proc ::NetPlacementUtility::Utils::getPlacementOrigin {} {
    set originX 0.0
    set originY 0.0
    
    # 1. 获取当前活跃页面对象
    set activePage [GetActivePage]
    
    # 2. 检查页面是否有效
    if {$activePage eq "NULL" || $activePage eq ""} {
        return [list $originX $originY]
    }
    
    # 3. 尝试获取点击坐标
    if {![catch {set clickPoint [GetLastMouseClickPointOnPage]} result]} {
        set originX [lindex $clickPoint 0]
        set originY [lindex $clickPoint 1]
    }
    
    return [list $originX $originY]
}

# --------------------------------------------------------------------------
# 核心功能：放置导线和网络标号 (Array 模式)
# API 参考: PlaceWire , PlaceNetAlias 
# --------------------------------------------------------------------------
proc ::NetPlacementUtility::Utils::placeNetsArray {prefix startNum endNum suffix} {
    set origin [getPlacementOrigin]
    set originX [lindex $origin 0]
    set originY [expr {[lindex $origin 1] - 0.1}] ;# 微调Y轴对齐格点

    set suffix [string toupper $suffix]  
    set maxStrLen [expr {[string length $startNum] > [string length $endNum] ? \
                  [string length $prefix$startNum$suffix] : [string length $prefix$endNum$suffix]}]
    
    # 计算导线长度 (基于字符长度估算 - 优化版)
    set wLenStart 0
    # 缩减系数 0.07 -> 0.05, 缓冲 0.4 -> 0.2
    set wLenEnd [expr {round(10 * ($maxStrLen * 0.05 + $wLenStart + 0.2)) * 0.1}]
    set count [expr {abs($startNum - $endNum) + 1}]
    
    for {set i 1} {$i <= $count} {incr i} {
        set yOffset [expr {$i / 10.0}]
        
        # 绘制导线
        PlaceWire [expr {$originX + $wLenStart}] [expr {$originY + $yOffset}] \
                  [expr {$originX + $wLenEnd}]   [expr {$originY + $yOffset}]
        
        # 计算当前编号
        if {$startNum >= $endNum} {
            set currentNum [expr {$startNum - $i + 1}]
        } else {
            set currentNum [expr {$startNum + $i - 1}]
        }
        set netName "${prefix}${currentNum}${suffix}"

        # 放置别名 (统一左对齐，移除 +0.2 偏移)
        PlaceNetAlias [expr {$originX + $wLenStart}] [expr {$originY + $yOffset}] $netName
    }
}

# --------------------------------------------------------------------------
# 核心功能：放置导线和网络标号 (List 模式)
# --------------------------------------------------------------------------
proc ::NetPlacementUtility::Utils::placeNetsList {prefix netNames suffix} {
    set origin [getPlacementOrigin]
    set originX [lindex $origin 0]
    set originY [expr {[lindex $origin 1] - 0.1}]
    
    set netNames [string toupper $netNames]
    set suffix [string toupper $suffix]
    
    # 计算最大长度
    set maxStrLen 0
    foreach net $netNames {
        set curLen [string length "${prefix}${net}${suffix}"]
        if {$curLen > $maxStrLen} { set maxStrLen $curLen }
    }
    
    set wLenStart 0
    # 缩减系数 0.07 -> 0.05, 缓冲 0.4 -> 0.2
    set wLenEnd [expr {round(10 * ($maxStrLen * 0.05 + $wLenStart + 0.2)) * 0.1}]
    
    set i 0
    foreach net $netNames {
        incr i
        if {$net eq ""} continue
        
        # 过滤非法网络名
        if {[regexp -nocase {^(NC[0-9]*|X|VCC|VDD|VSS|GND)$} $net]} continue

        set yOffset [expr {$i / 10.0}]
        PlaceWire [expr {$originX + $wLenStart}] [expr {$originY + $yOffset}] \
                  [expr {$originX + $wLenEnd}]   [expr {$originY + $yOffset}]
        PlaceNetAlias [expr {$originX + $wLenStart}] [expr {$originY + $yOffset}] \
                      "${prefix}${net}${suffix}"
    }
}

# --------------------------------------------------------------------------
# 核心功能：放置 Off-Page Connector
# API 参考: PlaceOffPage [cite: 1293]
# --------------------------------------------------------------------------
proc ::NetPlacementUtility::Utils::placeOffpageCommon {prefix startNum endNum suffix direction listMode netList} {
    variable ::NetPlacementUtility::options
    
    set origin [getPlacementOrigin]
    set originX [lindex $origin 0]
    set originY [expr {[lindex $origin 1] - 0.1}]
    
    set suffix [string toupper $suffix]
    set libPath $options(libPath)
    
    # 确定符号名称
    set symbolName [expr {$direction eq "in" ? "OFFPAGELEFT-L" : "OFFPAGELEFT-R"}]
    
    if {$listMode} {
        set items $netList
        set count [llength $items]
    } else {
        set count [expr {abs($startNum - $endNum) + 1}]
    }

    for {set i 1} {$i <= $count} {incr i} {
        set yOffset [expr {$i / 10.0}]
        
        if {$listMode} {
            set baseName [lindex $items [expr {$i - 1}]]
            if {$baseName eq ""} continue
             # 过滤非法网络名
            if {[regexp -nocase {^(NC[0-9]*|X|VCC|VDD|VSS|GND)$} $baseName]} continue
            set netName "${prefix}${baseName}${suffix}"
        } else {
            if {$startNum >= $endNum} {
                set currentNum [expr {$startNum - $i + 1}]
            } else {
                set currentNum [expr {$startNum + $i - 1}]
            }
            set netName "${prefix}${currentNum}${suffix}"
        }
        
        # 执行放置命令
        PlaceOffPage [expr {$originX}] [expr {$originY + $yOffset}] \
                     $libPath $symbolName $netName
                     
        # 针对 Output (Right) 方向调整 Name 属性位置
        # 用户反馈 Output 时名称太靠左，需要向右调整
        if {$direction eq "out"} {
            set lStatus [DboState]
            
            # 获取选中对象 (PlaceOffPage 通常会将新对象加入选择集)
            set objs [GetSelectedObjects]
            set numObjs [llength $objs]
            
            if {$numObjs > 0} {
                # 总是取最后一个对象 (假设新放置的对象被追加到列表末尾)
                # 这样即使选择集中有多个对象，也能处理最新的那个
                set obj [lindex $objs end]
                
                set nameCStr [DboTclHelper_sMakeCString "Name"]
                set pProp [$obj GetDisplayProp $nameCStr $lStatus]
                
                if {$pProp != "NULL"} {
                    # 计算长度
                    set curLen [string length "$netName"]
                    # 缩减系数 0.5, 缓冲 0
                    set locX [expr {int(-($curLen * 0.5 + 0)*10)}]

                    # 设置相对位置 (根据经验调整，OffPageLeft-R 原点在右侧尖端)
                    set loc [DboTclHelper_sMakeCPoint $locX 5]
                    $pProp SetLocation $loc
                }
                DboTclHelper_sDeleteCString $nameCStr
            }
            $lStatus -delete
        }
    }
}

# 包装函数
proc ::NetPlacementUtility::Utils::placeOffpageInArray {pre start end suf} {
    placeOffpageCommon $pre $start $end $suf "in" 0 {}
}
proc ::NetPlacementUtility::Utils::placeOffpageInList {pre list suf} {
    placeOffpageCommon $pre 0 0 $suf "in" 1 $list
}
proc ::NetPlacementUtility::Utils::placeOffpageOutArray {pre start end suf} {
    placeOffpageCommon $pre $start $end $suf "out" 0 {}
}
proc ::NetPlacementUtility::Utils::placeOffpageOutList {pre list suf} {
    placeOffpageCommon $pre 0 0 $suf "out" 1 $list
}

# --------------------------------------------------------------------------
# Database 操作：获取选中引脚名称
# 规则: 必须正确处理 DboState, CString, 和 Iterator 的内存释放 
# API 参考: GetSelectedObjects [cite: 1302], DboTclHelper_sMakeCString [cite: 1932]
# --------------------------------------------------------------------------
proc ::NetPlacementUtility::Utils::getSelectedPinNames {} {
    set pinNames {}
    
    # 1. 获取选中对象列表
    set selectedObjs [GetSelectedObjects]
    if {[llength $selectedObjs] == 0} { return $pinNames }

    # 2. 初始化状态和辅助对象
    set lStatus [DboState]
    set lPropValueCStr [DboTclHelper_sMakeCString]
    
    foreach lObj $selectedObjs {
        # 3. 关键修复：检查对象类型。我们只关心 DboPortInst (Type 22)
        # 注意：GetPinName 是 DboPortInst 的方法 [cite: 1716]
        # 如果选中的是 PartInst (Type 12), 调用 GetPinName 会崩溃或报错
        
        set objType [$lObj GetObjectType]
        
        # DboBaseObject::PORT_INST = 22 (需查阅 C++ 头文件或通过测试确认，这里用 catch 保护)
        if {[catch {$lObj GetPinName $lPropValueCStr} err] == 0} {
            lappend pinNames [DboTclHelper_sGetConstCharPtr $lPropValueCStr]
        }
    }
    
    # 4. 内存清理：非常重要！防止内存泄漏
    DboTclHelper_sDeleteCString $lPropValueCStr
    $lStatus -delete
    
    return $pinNames
}

proc ::NetPlacementUtility::Utils::placePinNetsArray {prefix suffix} {
    set pinNames [getSelectedPinNames]
    
    if {[llength $pinNames] == 0} {
        tk_messageBox -icon warning -title "提示" -message "请先选中一个或多个引脚(Pins)，不要选中元件整体。"
        return
    }
    
    # 复用列表放置逻辑
    placeNetsList $prefix $pinNames $suffix
}

#===============================================================================
# Callbacks: UI 事件处理
#===============================================================================
namespace eval ::NetPlacementUtility::Callbacks {
    namespace export executePlacement getPinNamesToList
}

proc ::NetPlacementUtility::Callbacks::executePlacement {} {
    variable ::NetPlacementUtility::options
    
    set opType $options(operationType)
    set genMethod $options(generationMethod)
    set dir $options(offpageDirection)
    
    # 从UI获取列表内容
    set netNamesUi [::NetPlacementUtility::UI::getNetnamesText]
    
    # 逻辑分发
    if {$opType eq "nets"} {
        if {$genMethod eq "array"} {
            ::NetPlacementUtility::Utils::placeNetsArray $options(prefix) $options(startNum) $options(endNum) $options(suffix)
        } else {
            if {[llength $netNamesUi] == 0} { tk_messageBox -message "列表为空"; return }
            ::NetPlacementUtility::Utils::placeNetsList $options(prefix) $netNamesUi $options(suffix)
        }
    } elseif {$opType eq "offpage"} {
        if {$dir eq "in"} {
            if {$genMethod eq "array"} {
                ::NetPlacementUtility::Utils::placeOffpageInArray $options(prefix) $options(startNum) $options(endNum) $options(suffix)
            } else {
                if {[llength $netNamesUi] == 0} { tk_messageBox -message "列表为空"; return }
                ::NetPlacementUtility::Utils::placeOffpageInList $options(prefix) $netNamesUi $options(suffix)
            }
        } else {
            if {$genMethod eq "array"} {
                ::NetPlacementUtility::Utils::placeOffpageOutArray $options(prefix) $options(startNum) $options(endNum) $options(suffix)
            } else {
                if {[llength $netNamesUi] == 0} { tk_messageBox -message "列表为空"; return }
                ::NetPlacementUtility::Utils::placeOffpageOutList $options(prefix) $netNamesUi $options(suffix)
            }
        }
    }
}

proc ::NetPlacementUtility::Callbacks::getPinNamesToList {} {
    set pinNames [::NetPlacementUtility::Utils::getSelectedPinNames]
    if {[llength $pinNames] > 0} {
        # 自动切换UI状态
        set ::NetPlacementUtility::options(operationType) "nets"
        set ::NetPlacementUtility::options(generationMethod) "list"
        ::NetPlacementUtility::UI::updateUiLayout
        
        ::NetPlacementUtility::UI::setNetnamesText $pinNames
    } else {
        tk_messageBox -icon warning -title "提示" -message "未检测到选中的引脚对象。\n请确保使用的是 Alt+拖动 选择引脚，\n而不是选中整个元件。"
    }
}

#===============================================================================
# UI: 界面构建
#===============================================================================
namespace eval ::NetPlacementUtility::UI {
    namespace export createWindow updateUiLayout
}

proc ::NetPlacementUtility::UI::createWindow {} {
    variable ::NetPlacementUtility::widgets
    variable ::NetPlacementUtility::options
    
    if {[winfo exists .netplacement]} {
        destroy .netplacement
        # 状态重置
        set ::NetPlacementUtility::state(windowCreated) 0
    }
    
    # 定义字体
    set uiFont "微软雅黑 9" 
    
    set win [toplevel .netplacement]
    wm title $win "网络放置工具"
    wm attributes $win -topmost 1
    wm protocol $win WM_DELETE_WINDOW [list ::NetPlacementUtility::hide]
    set widgets(mainWindow) $win

    set main [frame $win.main -padx 10 -pady 10]
    pack $main -fill both -expand true
    
    # 1. 操作类型
    set frType [labelframe $main.type -text "操作模式" -padx 5 -pady 5 -font $uiFont]
    pack $frType -fill x -pady 5
    
    foreach {lbl val} {"网络标号" nets "Off-Page" offpage} {
        radiobutton $frType.rb_$val -text $lbl -variable ::NetPlacementUtility::options(operationType) \
            -value $val -command ::NetPlacementUtility::UI::updateUiLayout -font $uiFont
        pack $frType.rb_$val -side left -padx 5
    }

    # 2. 生成方式与参数
    set frParam [labelframe $main.param -text "参数设置" -padx 5 -pady 5 -font $uiFont]
    pack $frParam -fill x -pady 5
    set widgets(frParam) $frParam

    # 2.1 方法选择
    set frMethod [frame $frParam.method]
    set widgets(frMethod) $frMethod
    radiobutton $frMethod.rb_arr -text "阵列 (1,2...)" -variable ::NetPlacementUtility::options(generationMethod) \
        -value "array" -command ::NetPlacementUtility::UI::updateUiLayout -font $uiFont
    radiobutton $frMethod.rb_lst -text "列表 (Names)" -variable ::NetPlacementUtility::options(generationMethod) \
        -value "list" -command ::NetPlacementUtility::UI::updateUiLayout -font $uiFont
    pack $frMethod.rb_arr $frMethod.rb_lst -side left -padx 5

    # 2.2 阵列输入
    set frArray [frame $frParam.array]
    set widgets(frArray) $frArray
    label $frArray.l1 -text "Start:" -font $uiFont
    entry $frArray.e1 -textvariable ::NetPlacementUtility::options(startNum) -width 5 -font $uiFont
    label $frArray.l2 -text "End:" -font $uiFont
    entry $frArray.e2 -textvariable ::NetPlacementUtility::options(endNum) -width 5 -font $uiFont
    pack $frArray.l1 $frArray.e1 $frArray.l2 $frArray.e2 -side left -padx 2

    # 2.3 列表输入
    set frList [frame $frParam.list]
    set widgets(frList) $frList
    label $frList.tip -text "分隔符支持：回车、空格、逗号(,)、分号(;)" -fg "#666666" -font "微软雅黑 8" -anchor w
    pack $frList.tip -side top -fill x -pady {0 2}
    text $frList.txt -width 30 -height 8 -font $uiFont
    scrollbar $frList.sb -command "$frList.txt yview"
    $frList.txt configure -yscrollcommand "$frList.sb set"
    set widgets(netnamesText) $frList.txt
    pack $frList.sb -side right -fill y
    pack $frList.txt -side left -fill both -expand true
    
    # 2.4 前缀后缀
    set frFix [frame $frParam.fix -pady 5]
    set widgets(frFix) $frFix
    label $frFix.lpre -text "Prefix:" -font $uiFont
    entry $frFix.epre -textvariable ::NetPlacementUtility::options(prefix) -width 8 -font $uiFont
    label $frFix.lsuf -text "Suffix:" -font $uiFont
    entry $frFix.esuf -textvariable ::NetPlacementUtility::options(suffix) -width 8 -font $uiFont
    pack $frFix.lpre $frFix.epre $frFix.lsuf $frFix.esuf -side left -padx 2

    # 3. Off-Page 专属设置
    set frOff [labelframe $main.off -text "Off-Page Config" -padx 5 -pady 5 -font $uiFont]
    set widgets(frOff) $frOff
    radiobutton $frOff.in -text "Input (Left)" -variable ::NetPlacementUtility::options(offpageDirection) -value "in" -font $uiFont
    radiobutton $frOff.out -text "Output (Right)" -variable ::NetPlacementUtility::options(offpageDirection) -value "out" -font $uiFont
    entry $frOff.lib -textvariable ::NetPlacementUtility::options(libPath) -font $uiFont
    pack $frOff.in $frOff.out -side left
    pack $frOff.lib -fill x -pady 5

    # 4. 底部按钮
    set frBtn [frame $main.btn -pady 5]
    pack $frBtn -fill x
    
    # 引脚获取按钮 (仅在List模式或引脚模式显示)
    button $frBtn.getPin -text "获取选中引脚名" -command ::NetPlacementUtility::Callbacks::getPinNamesToList -font $uiFont
    set widgets(btnGetPin) $frBtn.getPin

    button $frBtn.run -text "执行放置" -bg "#DDDDDD" -command ::NetPlacementUtility::Callbacks::executePlacement -font $uiFont
    set widgets(btnRun) $frBtn.run

    updateUiLayout
}

proc ::NetPlacementUtility::UI::updateUiLayout {} {
    variable ::NetPlacementUtility::options
    variable ::NetPlacementUtility::widgets
    
    set type $options(operationType)
    set method $options(generationMethod)

    # 清理显示
    pack forget $widgets(frMethod)
    pack forget $widgets(frArray)
    pack forget $widgets(frList)
    pack forget $widgets(frFix)
    pack forget $widgets(frOff)
    pack forget $widgets(btnGetPin)
    $widgets(btnRun) configure -state normal

    # 直接执行原本 else 分支的内容
    pack $widgets(frMethod) -fill x
    pack $widgets(frFix) -fill x
    
    if {$method eq "array"} {
        pack $widgets(frArray) -fill x
    } else {
        pack $widgets(frList) -fill both -expand true
        # 注意：这保留了在“列表”模式下获取引脚的功能
        pack $widgets(btnGetPin) -side left
    }

    if {$type eq "offpage"} {
        pack $widgets(frOff) -fill x
    }
    
    pack $widgets(btnRun) -side right
}

proc ::NetPlacementUtility::UI::getNetnamesText {} {
    variable ::NetPlacementUtility::widgets
    set content [$widgets(netnamesText) get 1.0 "end-1c"]
    set cleanList {}

    # 使用正则表达式将：换行(\n)、回车(\r)、逗号(,)、分号(;)、制表符(\t) 统一替换为空格
    regsub -all {[\n\r,;\t]+} $content " " content
    
    # Tcl中以空格分隔的字符串可以直接作为列表遍历
    foreach item $content {
        set item [string trim $item]
        # 过滤空元素
        if {$item ne ""} { lappend cleanList $item }
    }

    return $cleanList
}

proc ::NetPlacementUtility::UI::setNetnamesText {listData} {
    variable ::NetPlacementUtility::widgets
    $widgets(netnamesText) delete 1.0 end
    $widgets(netnamesText) insert end [join $listData "\n"]
}

#===============================================================================
# Main: 入口
#===============================================================================
proc ::NetPlacementUtility::show {} {
    variable state
    variable widgets
    if {![winfo exists .netplacement]} {
        ::NetPlacementUtility::UI::createWindow
        set state(windowCreated) 1
    }
    wm deiconify .netplacement
    raise .netplacement
}

proc ::NetPlacementUtility::hide {} {
    if {[winfo exists .netplacement]} {
        wm withdraw .netplacement
    }
}

#puts "Run '::NetPlacementUtility::show' to start."