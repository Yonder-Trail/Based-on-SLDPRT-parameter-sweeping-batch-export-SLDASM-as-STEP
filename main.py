##########################################################Based-on-SLDPRT-parameter-sweeping-batch-export-SLDASM-as-STEP##########################################################
#本项目本项目依托SolidWorks COM API与Python实现自动化建模功能，可通过修改零件全局变量驱动装配体参数化重建，并完成 STEP文件的批量导出。
#逻辑：用户使用前需要确保带有全局参数的零件文件以及与其相关的装配体文件存在。
#本代码运行逻辑：首先改变零件文件全局变量，再重建零件文件，之后基于重建的零件文件重建装配体文件，最后导出为STEP文件，整体为批量操作，特别地，代码运行中配合关系是保持不变的。
import os
import gc
import math
import time
import itertools
import pythoncom
import subprocess
import numpy as np
import win32com.client
from win32com.client import VARIANT
from datetime import datetime, timedelta
#Based-on-SLDPRT-parameter-sweeping-batch-export-SLDASM-as-STEP
def SLDASM_to_STEP_NewRebuild(Global_Parameter,Global_Parameter_list,SLDPRT_Filesname,SLDPRT_Dir,SLDASM_files_Dir,Output_STEP_Dir,Componentindex,Planindex):
    #创建输出目录
    def Create_Directory(Dir_path):
        if not os.path.exists(Dir_path):
            os.makedirs(Dir_path)
    #获取当前进程信息
    def Get_Progcess_Information(start_time,Countindex,SUM,Planindex):
        global SW_app
        Elapsed_Time=time.time()-start_time
        Speed=Elapsed_Time/Countindex
        Remaining_Sec=(SUM-Countindex)*Speed
        Remaining_Time=timedelta(seconds=int(Remaining_Sec))
        Eta_Time=datetime.now()+timedelta(seconds=Remaining_Sec)
        Eta_str=Eta_Time.strftime("%Y-%m-%d %H:%M:%S")
        Elapsed_str=str(timedelta(seconds=int(Elapsed_Time)))
        Progress=(Countindex/SUM)*100
        if Countindex%200==0:
            #重新启动SolidWorks软件，释放内存空间
            SW_app.ExitApp()
            del SW_app
            gc.collect()
            subprocess.call('taskkill /f /im SLDWORKS.exe', shell=True)
            SW_app=win32com.client.Dispatch("SldWorks.Application")
            SW_app.Visible=False
        print(f"[OK] 当前计划 {Planindex} | 完成 {(Countindex)}/{SUM} | 进度：{Progress:.3f}%    [Time] 已用时间：{Elapsed_str} | 剩余时间：{Remaining_Time} | 预计完成：{Eta_str}")

    #重建装配体
    def Rebuild_SLDASM_to_STEP(Global_Parameter,Global_Parameter_list,SLDPRT_Dir,SLDASM_Target_Files,Output_STEP_Incomplete_Dir,Planindex):
        global SW_app
        #全局变量的各参量组合
        All_Parameter_Combinations=[]
        for Global_Parameter_index in range(len(Global_Parameter)):
            Lower_Limit_Value,Upper_Limit_Value,Step_Value=Global_Parameter_list[Global_Parameter_index]
            All_Parameter_Combinations.append(np.arange(Lower_Limit_Value,Upper_Limit_Value,Step_Value))
        # 启动SolidWorks软件
        SW_app=win32com.client.Dispatch("SldWorks.Application")
        SW_app.Visible=False
        #定义错误与警告变量   
        Errors=VARIANT(pythoncom.VT_BYREF|pythoncom.VT_I4,0)
        Warnings=VARIANT(pythoncom.VT_BYREF|pythoncom.VT_I4,0)
        #总数
        SUM=math.prod([len(Parameter_Arr) for Parameter_Arr in All_Parameter_Combinations])*len(SLDASM_Target_Files)
        start_time=time.time()#开始时间
        Countindex=1#当前进度
        for _,Parameter_Combination in enumerate(itertools.product(*All_Parameter_Combinations)):
            #零件全局参量修改
            SW_SLDPRT_Doc=SW_app.OpenDoc6(SLDPRT_Dir,1,1,"",Errors,Warnings)#打开零件
            SW_model=SW_app.ActiveDoc
            Eq_Mgr=SW_model.GetEquationMgr
            Mgr_Count=Eq_Mgr.GetCount#方程数量
            for Mgr_Count_index in range(Mgr_Count):
                try:
                    Is_Global=Eq_Mgr.GlobalVariable(Mgr_Count_index)
                except:
                    Is_Global=False
                if Is_Global:
                    Eq_text=Eq_Mgr.Equation(Mgr_Count_index)
                    for Global_Parameter_index,Global_Parameter_Text in enumerate(Global_Parameter):
                        if Global_Parameter_Text in Eq_text:
                            Eq_Mgr.Equation(Mgr_Count_index,f'"{Global_Parameter_Text}"= {Parameter_Combination[Global_Parameter_index]}mm')#修改全局变量
                        else:
                            pass
            SW_model.ForceRebuild3(True)#强制重建
            SW_SLDPRT_Doc.Save()#保存零件
            SW_app.CloseDoc(f"{os.path.splitext(os.path.basename(SLDPRT_Dir))[0]}.SLDPRT")#关闭零件
            del Eq_Mgr,SW_model,SW_SLDPRT_Doc
            gc.collect()#回收内存
            #导出为STEP文件
            for SLDASM_Target_File_Index,SLDASM_Target_File in enumerate(SLDASM_Target_Files):
                Output_STEP_Complete_Dir=Output_STEP_Incomplete_Dir[SLDASM_Target_File_Index]+f"_{Countindex}.step"
                SW_SLDASM_Doc=SW_app.OpenDoc6(SLDASM_Target_File,2,1,"",Errors,Warnings)#打开装配体(数字2表示装配体文件)
                SW_models=SW_app.ActiveDoc
                SW_models.ForceRebuild3(True)#强制重建
                SW_SLDASM_Doc.SaveAs4(Output_STEP_Complete_Dir,2,1,Errors,Warnings)#保存为STEP文件(数字2表示STEP格式)
                SW_app.CloseDoc(f"{os.path.splitext(os.path.basename(SLDASM_Target_File))[0]}.SLDASM")
                del SW_models,SW_SLDASM_Doc
                gc.collect()#回收内存
                Get_Progcess_Information(start_time,Countindex,SUM,Planindex)
                Countindex=Countindex+1
        SW_app.ExitApp()               
    #创建输出目录
    Create_Directory(Output_STEP_Dir) 
    #目标装配体路径地址完整声明
    SLDASM_Target_Files=[]
    for SLDPRT_Filesname_target in SLDPRT_Filesname:
        SLDASM_Target_Files.append(os.path.join(SLDASM_files_Dir,f"{SLDPRT_Filesname_target}.SLDASM").replace('\\', '/'))
    #装配体STEP输出路径地址非完整声明，后续需进行完整声明，在Rebuild_SLDASM_to_STEP()函数中见Output_STEP_Complete_Dir变量
    Output_STEP_Incomplete_Dir=[]
    for SLDPRT_Filesname_target in SLDASM_Target_Files:
        File_name=os.path.splitext(os.path.basename(SLDPRT_Filesname_target))[0]
        Output_STEP_Incomplete_Dir.append(os.path.join(Output_STEP_Dir, f"{File_name}_{Componentindex}_").replace('\\', '/'))
    #重建装配体并转STEP
    Rebuild_SLDASM_to_STEP(Global_Parameter,Global_Parameter_list,SLDPRT_Dir,SLDASM_Target_Files,Output_STEP_Incomplete_Dir,Planindex)
    
if __name__ == "__main__":
    #参数设置
    Planindex=1#当前计划索引，用于区分不同的计划
    Componentindex="component1"#表示当前零件字符串索引
    Global_Parameter=["L1","L2","L3"]#目标零件的全局变量参数,可多个变量
    Global_Parameter_list=[[1000,3000,100],[1000,3000,100],[1000,3000,100]]#全局变量参数变化范围设置，单位为mm，每个三个参数[Lower_Limit_Value,Upper_Limit_Value,Step_Value]为一组，分别表示为全局变量的下限、上限以及步长
    SLDPRT_Filesname=["Assembly1","Assembly2"]#与目标零件相关的装配体文件名列表，可多个变量
    SLDPRT_Dir=f""#目标零件路径地址
    SLDASM_files_Dir=r""#装配体路径地址
    Output_STEP_Dir=r""#重建装配体后转STEP的输出路径地址
    #Based-on-SLDPRT-parameter-sweeping-batch-export-SLDASM-as-STEP
    SLDASM_to_STEP_NewRebuild(Global_Parameter,Global_Parameter_list,SLDPRT_Filesname,SLDPRT_Dir,SLDASM_files_Dir,Output_STEP_Dir,Componentindex,Planindex)
