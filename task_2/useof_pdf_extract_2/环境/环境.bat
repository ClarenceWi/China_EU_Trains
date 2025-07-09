echo 正在配置pip使用阿里云镜像源...
echo.

REM 步骤1: 创建pip配置文件目录
if not exist "%APPDATA%\pip" (
    mkdir "%APPDATA%\pip"
    echo 已创建pip配置目录: %APPDATA%\pip
)

REM 步骤2: 创建或更新pip.ini配置文件
(
    echo [global]
    echo index-url = https://mirrors.aliyun.com/pypi/simple/
    echo trusted-host = mirrors.aliyun.com
    echo timeout = 6000
) > "%APPDATA%\pip\pip.ini"

echo 配置文件已更新: %APPDATA%\pip\pip.ini
echo.
echo 配置文件内容:
type "%APPDATA%\pip\pip.ini"
echo.

REM 步骤3: 验证配置
echo 验证pip配置...
pip config list
echo.

echo 配置完成

pip install pdfplumber 
pip install pandas 
pip install openpyxl
