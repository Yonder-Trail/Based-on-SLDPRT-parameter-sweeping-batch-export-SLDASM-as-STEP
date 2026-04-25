# Based-on-SLDPRT-parameter-sweeping-batch-export-SLDASM-as-STEP

本项目依托 SolidWorks COM API 与 Python 实现自动化建模功能，可通过修改零件全局变量驱动装配体参数化重建，并完成 STEP 文件的批量导出。

Leverages SolidWorks COM API and Python for automated modeling. Enables assembly parametric reconstruction through modification of part global variables, with full support for batch STEP file export.

使用前注意事项：

(1) main.py为主程序;

(2) 使用前请先确保存在零件文件(.SLDPRT)与该零件捆绑的装配体文件(.SLDASM)存在目标文件夹，其中零件的目标文件夹在变量SLDPRT_Dir上，装配体的目标文件夹在变量SLDASM_files_Dir上;

(3) 使用前应当将变量Planindex、Componentindex、Global_Parameter、Global_Parameter_list、SLDPRT_Filesname、SLDPRT_Dir、SLDASM_files_Dir、Output_STEP_Dir重新根据自己的需求进行填写。

(4) 本代码运行逻辑：首先改变零件文件全局变量，再重建零件文件，之后基于重建的零件文件重建装配体文件，最后导出为STEP文件，整体为批量操作。


Precautions Before Use:

(1) main.py is the main program.

(2) Before use, ensure that part files (.SLDPRT) and their matched assembly files (.SLDASM) exist in the target folders. The storage path for part files is defined by the variable SLDPRT_Dir, and the storage path for assembly files is specified by the variable SLDASM_files_Dir.

(3) Modify and configure the following variables according to actual requirements before operation:Planindex, Componentindex, Global_Parameter, Global_Parameter_list, SLDPRT_Filesname, SLDPRT_Dir, SLDASM_files_Dir and Output_STEP_Dir.

(4) Code Execution Logic: First, modify the global variables of the part file, then rebuild the part file. Next, rebuild the assembly file based on the updated part file. Finally, export the assembly as a STEP file. The entire process is a batch operation.
