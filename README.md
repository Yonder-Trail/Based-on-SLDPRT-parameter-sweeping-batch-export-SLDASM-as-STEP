# Based-on-SLDPRT-parameter-sweeping-batch-export-SLDASM-as-STEP

本项目依托 SolidWorks COM API 与 Python 实现自动化建模功能，可通过修改零件全局变量驱动装配体参数化重建，并完成 STEP 文件的批量导出。

Leverages SolidWorks COM API and Python for automated modeling. Enables assembly parametric reconstruction through modification of part global variables, with full support for batch STEP file export.

使用前注意事项：
(1) main.py为主程序;
(2) 使用前请先确保存在零件文件(.SLDPRT)与该零件捆绑的装配体文件(.SLDASM)存在目标文件夹，其中零件的目标文件夹在变量SLDPRT_Dir上，装配体的目标文件夹在变量SLDASM_files_Dir上;
(3) 使用前应当将变量Planindex、Componentindex、Global_Parameter、Global_Parameter_list、SLDPRT_Filesname、SLDPRT_Dir、SLDASM_files_Dir、Output_STEP_Dir重新根据自己的需求进行填写。
