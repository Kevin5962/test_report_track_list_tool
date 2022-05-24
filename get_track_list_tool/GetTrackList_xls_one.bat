rem 1.在本地路径新建一个文件夹，把脚本和bat文件放进去，修改文件名为你想要分析的测试报告
rem 2.把需要分析的测试报告文件，打开并另存为xls格式，因为本脚本只能处理.xls格式的excel文件，切记：是打开文件并另存为的方式，不能直接修改后缀
rem 3.本脚本的输入参数只能有一个
echo -----生成 track list-------
C:/Python38/python.exe get_track_list_xlrd_one.py  0422-BMW_IDC_Function_Test_Summary_ENG_2022_05_18_16_14_00.xls
rem 0422-BMW_IDC_Function_Test_Summary_ENG_2022_05_18_16_14_00.xls
rem BMW_IDC_Function_Test_Summary_MNC_2022_05_18_11_47_24.xls
rem 0511版本-BMW_IDC_Function_Test_Summary_MNC_2022_05_12_09_04_30.xls   
rem BMW_IDC_Function_Test_Summary_MNC_2022_05_13_17_10_21.xls
rem 更新0427版本-BMW_IDC_Function_Test_Summary_MNC_2022_04_29.xls
cmd.exe