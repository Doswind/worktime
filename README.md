### 使用说明

1. 二进制包使用说明
   - `worktime_windows.exe` ： 可以 `windows7` 及以上环境上使用。
   - `worktime_linux`： 可在 `Linux` 环境上使用
   
2. 源码包使用说明

   此工具使用 `python3 `开发，所以需要 `python3` 及相关依赖项。
  - 源码文件说明
    - `requirement.sh`：运行环境安装脚本
    - `worktime.py`：工具脚本
  - 运行环境安装
      可执行命令如下：

    ```bash
    sudo apt -y install python3 python3-pip python3-tk
    sudo pip3 install openpyxl
    ```
    也可以使用封装好的脚本,执行后输入操作系统用户密码完成安装。
    ```bash
    bash requirement.sh
    ```
  - 工具使用方法
      执行如下命令，弹出应用界面，查看`帮助`进行操作。
    ```bash
    python3 worktime.py
    ```

  

