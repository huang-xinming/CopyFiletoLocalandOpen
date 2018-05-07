#################################################################
# File      : CopyFiletoLocalandOpen
# Date      : 2017.10.22
# Created by: HXM
# Purpose   : Hyperworks调用，从远端将文件copy至Hyperworks当前模型所在目录
#			  打开copy的文件
#################################################################
###############################################################
## 定义变量
###############################################################
catch {namespace delete ::NVH::}
namespace eval ::NVH:: {
	#运行环境
	variable baseDir;#tcl脚本运行目录
	set baseDir [file dir [info script]];
	#被copy的文件信息
	variable filename;#文件名称 
	set filename "Tow_DataPostProcessing.xlsm";
	variable sourcefile;
	set sourcefile "$baseDir/$filename";
}
###############################################################
## 获取当前模型所在目录
###############################################################
proc ::NVH::GetDestination {} {
	variable destdir;
	set currentfile [hm_info currentfile]
	set destdir [file dir $currentfile]
	return $destdir
}
###############################################################
## 搜索当前目录下是否存在类似文件
###############################################################
proc ::NVH::CheckFile {filename} {
	set destdir [::NVH::GetDestination]
	if {[catch {glob -directory "$destdir" $filename}]==1} {
		set filestatus 1;#没有文件
		return $filestatus;
	} elseif {[catch {glob -directory "$destdir" $filename}]==0} {
		set filestatus [glob -directory "$destdir" $filename]
		return [file normalize $filestatus];
	}
}
###############################################################
## 从远端copy文件
###############################################################
proc ::NVH::CopyFile {} {
	variable sourcefile;
	set destdir [::NVH::GetDestination]
	catch {file copy "$sourcefile" "$destdir"}
}
###############################################################
## Main: 检查当前文件目录是否有相应文件，如果没有则copy并打开，如果有则直接打开
###############################################################
proc ::NVH::Main {} {
	variable filename;
	variable sourcefile;
	
	set flag [::NVH::CheckFile $filename]
	if {$flag==1} {
		::NVH::CopyFile
		set copiedfile [::NVH::CheckFile $filename]
		exec cmd.exe /c $copiedfile &
	} else {
		set answer [tk_messageBox -title "NVH Tools" -icon info -message "$flag exists in current directory, click yes to open existing file, click no to overwrite." \
		-type yesnocancel]
		switch $answer {
			cancel {
				exit
			}
			yes {
				exec cmd.exe /c $flag &
				return 0
			}
			no {
				if {[catch {file delete $flag}]==1} {
					tk_messageBox -title "NVH Tools" -icon error -message "$flag can not be Deleted,Please Check" \
					-type ok
					exit
				} else {
					::NVH::CopyFile
					set copiedfile [::NVH::CheckFile $filename]
					exec cmd.exe /c $copiedfile &
					return 0
				}
			}
		}
	}
}
::NVH::Main